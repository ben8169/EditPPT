import os
import re
import time
import json
import base64
import shutil
from copy import deepcopy


import traceback

from editppt.tools.tools import *
from editppt.utils.llm_client import call_llm, call_llm_gemini
from editppt.prompts import *
from editppt.utils.logger_manual import log_path
from editppt.utils.utils import parse_active_slide_objects

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
SCHEMA_PATH = BASE_DIR / "tools" / "tools_schema.json"

with open(SCHEMA_PATH, "r", encoding="utf-8") as f:
    TOOLS_SCHEMA = json.load(f)  



class EditAgent:
    def __init__(self, container, model: str):
        self.container = container
        self.model = model

        self.backup_path = os.path.abspath(self.container.prs.Name.replace(".pptx", "_.pptx"))
        if os.path.exists(self.backup_path):
             os.remove(self.backup_path)
        self.container.prs.SaveAs(self.backup_path)

        self.messages = []

    def run(self, task: dict, parser: object, vision_validator_agent: object):
        feedback = []
        max_retries = 5
        retry_count = 0

        print()

        while retry_count < max_retries:
            retry_count += 1
            payload_message = []
            page_number = task.get("page number")
            if not page_number:
                continue

            description = task.get("description", "")
            action = task.get("action", "")
            detailed_contents = task.get("contents", "")

            contents = parser.process(page_number)
            
            # system/user message 구성
            payload_message.append({
                "role": "system",
                "content": create_edit_agent_system_prompt(contents)
            })
            payload_message.append({
                "role": "user",
                "content": create_edit_agent_user_prompt(
                    page_number, description, action, detailed_contents, feedback
                )
            })

            with open(
                log_path(f"agent_payload_message_{page_number}.txt"),
                "w",
                encoding="utf-8"
            ) as f:
                for i, msg in enumerate(payload_message, 1):
                    f.write(f"[{i}]\n")
                    f.write(f"role: {msg.get('role')}\n")

                    content = msg.get("content")
                    if isinstance(content, list):
                        for block in content:
                            if block.get("type") == "text":
                                f.write(block.get("text", ""))
                                f.write("\n")
                    else:
                        f.write(str(content) + "\n")
                    f.write("\n" + "-" * 50 + "\n\n")

            response = call_llm(
                model=self.model,
                messages=payload_message,
                tools=TOOLS_SCHEMA,
                tool_choice="auto"
            )

            response_dict = response.model_dump()  

            with open(
                log_path(f"agent_toolcall_response_{page_number}.json"),
                "w",
                encoding="utf-8"
            ) as f:
                json.dump(response_dict, f, ensure_ascii=False, indent=4)

            # Tool Call Parsing
            tool_calls = []
            for item in response.output:
                if item.type == "function_call":
                    tool_calls.append({
                        "name": item.name,
                        "arguments": json.loads(item.arguments) if isinstance(item.arguments, str) else item.arguments,
                        "call_id": item.call_id
                    })

            # LLM 메시지 로깅
            with open(log_path(f"agent_Message_retry_{retry_count}.json"), "w", encoding="utf-8") as f:
                json.dump(self.messages, f, ensure_ascii=False, indent=4)

            #Tool Fail 시 fallback
            failed_tool_name = None
            failed_tool_args = None
            tool_error_reason = None
            
            
            # Tool 실행 
            for tool_call in tool_calls:
                function_name = tool_call["name"]
                function_args = tool_call["arguments"]

                logger.info(f"Tool Call: {function_name}({function_args})")

                if function_name == "set_text_style_preserve_runs":
                    function_args["slide_json"] = contents
                    for key in ["bold", "italic", "underline", "font_name", "font_size"]:
                        if key in function_args and function_args[key] is False:
                            del function_args[key]

                if function_name == "replace_shape_text":
                    function_args["slide_json"] = contents
                    function_args["agent_request"] = description, action, detailed_contents

                try:
                    self._execute_tool(function_name, function_args)

                except Exception as e:
                    failed_tool_name = function_name
                    failed_tool_args = deepcopy(function_args) 
                    tool_error_reason = f"{function_name} failed: {e}"
                    break

            if tool_error_reason:
                executed_calls_str = ", ".join(
                    [f"{tc['name']}({tc['arguments']})" for tc in tool_calls]
                )

                feedback.append(
                    f"""Tool Execution Failed
                    - Failed Tool: {failed_tool_name}
                    - Failed Args: {json.dumps(failed_tool_args, ensure_ascii=False)}
                    - Error: {tool_error_reason}
                    - Planned Tools: [{executed_calls_str}]
                    """
                )
                self._rollback_ppt("tool", tool_error_reason)
                continue
    

            # Validation 1 (Text/Logic)
            valid, reason = parser.update_after_edit(
                text_validation=True,
                model=self.model,
                page_number=page_number,
                description=description,
                action=action,
                detailed_contents=detailed_contents,
                used_tools = tool_calls
            )

            if valid:
                # Validation 2 (Vision)
                if vision_validator_agent is not None:
                    valid, reason = vision_validator_agent.process(
                        page_number=page_number,
                        agent_request=action,
                        parsed_contents=contents,
                        used_tools=tool_calls)
                    
                    if valid:
                        # [최종 성공]
                        parser.edit_history.setdefault(page_number, []).append(deepcopy(parser.database.get(page_number, None)))
                        parser.database[page_number] = parse_active_slide_objects(page_number, self.container.prs)

                        with open(log_path("parser_Edithistory.json"), "w", encoding="utf-8") as f:
                            json.dump(parser.edit_history, f, ensure_ascii=False, indent=4)
                        with open(log_path("parser_Database.json"), "w", encoding="utf-8") as f:
                            json.dump(parser.database, f, ensure_ascii=False, indent=4)

                        
                        # 성공 지점을 백업으로 갱신
                        self.container.prs.SaveCopyAs(self.backup_path)
                        break
                    else:
                        # Vision 실패 시 롤백 및 재시도 준비
                        executed_calls_str = ", ".join([f"{tc['name']}({tc['arguments']})" for tc in tool_calls])
                        feedback.append(f"Retry {retry_count} Vision Fail: {reason} | Tools: [{executed_calls_str}]")
                        self._rollback_ppt("vision",reason)
                        continue
            else:
                # Text/Logic 실패 시 롤백 및 재시도 준비
                executed_calls_str = ", ".join([f"{tc['name']}({tc['arguments']})" for tc in tool_calls])
                feedback.append(f"Retry {retry_count} Text Fail: {reason} | Tools: [{executed_calls_str}]")
                self._rollback_ppt("text", reason)
                continue

            # Feedback 파일 갱신
            with open(log_path(f"agent_Feedback_{page_number}.json"), "w", encoding="utf-8") as f:
                json.dump(feedback, f, ensure_ascii=False, indent=4)

    def _rollback_ppt(self, type, reason):
        """
        Close edited PPT and roll back to the last backup
        """
        logger.warning(f"{type} Feedback: {reason}")
        ppt_app = self.container.prs.Application
        self.container.prs.Close()
        time.sleep(0.5)
        self.container.prs = ppt_app.Presentations.Open(os.path.abspath(self.backup_path))

    def _execute_tool(self, name, args):
        if name not in FUNCTION_MAP:
            return f"Error: Tool '{name}' not found."
        try:
            return FUNCTION_MAP[name](self.container.prs, **args)
        except Exception as e:
            logger.error(f"Execution Error: {e}")
            logger.error(traceback.format_exc())
            return f"Error: {str(e)}"







class VisionValidatorAgent:
    def __init__(self, container, model: str):
        self.container = container
        self.model = model

        self.base_dir = Path(__file__).resolve().parents[1]

        self.output_dir = self.base_dir / ".SlideScreenshots"
        if self.output_dir.exists():
            shutil.rmtree(self.output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    @classmethod
    def create(cls, activate_valid: bool, container, model: str):
        if not activate_valid:
            return None
        return cls(container, model)

    def process(self, page_number, agent_request, parsed_contents, used_tools):
        """
        1. Extract slide PNG
        2. Call Vision LLM
        3. Parse and validate JSON result
        """

        # --- Slide Export ---
        slide = self.container.prs.Slides(page_number)

        image_path = self.output_dir / f"slide_{page_number}.png"
        slide.Export(str(image_path), "PNG")

        with open(image_path, "rb") as image_file:
            encoded_image = base64.b64encode(image_file.read()).decode("utf-8")

        # --- LLM Call ---
        if self.model.startswith("gpt"):
            messages_for_validation = [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "input_text",
                            "text": create_vision_validator_agent_system_prompt(
                                agent_request,
                                parsed_contents,
                                used_tools,
                            ),
                        },
                        {
                            "type": "input_image",
                            "image_url": f"data:image/png;base64,{encoded_image}",
                        },
                    ],
                }
            ]

            response = call_llm(
                model=self.model,
                messages=messages_for_validation,
            )

        elif self.model.startswith("gemini"):
            with open(image_path, "rb") as f:
                img_bytes = f.read()

            system_prompt = create_vision_validator_agent_system_prompt(
                agent_request,
                parsed_contents,
                used_tools,
            )

            response = call_llm_gemini(
                model=self.model,
                messages=system_prompt,
                image=img_bytes,
            )

        else:
            raise ValueError(f"Unsupported model for vision validation: {self.model}")

        response_text = (response.output_text or "").strip()
        print(response_text)

        # --- JSON Extraction ---
        json_match = re.search(r"\{.*\}", response_text, re.DOTALL)
        if not json_match:
            # fallback: "No issue" 텍스트 대응
            if "no" in response_text.lower()[:20]:
                return True, None
            return False, f"Invalid vision response format: {response_text}"

        try:
            data = json.loads(json_match.group())
            has_issues = data.get("HasCriticalIssues", "No")

            if has_issues == "No":
                logger.info("Vision validation passed.")
                return True, None

            # --- Issue Parsing ---
            issues = data.get("Issues", [])
            if not issues and "TechnicalDiagnosis" in data:
                issues = [data]

            reason_list = []
            for idx, issue in enumerate(issues):
                diag = issue.get("TechnicalDiagnosis", "Unknown issue")
                fix = issue.get("ActionableFix", "")
                reason_list.append(
                    f"Issue {idx + 1}: {diag} (Suggest: {fix})"
                )

            reason = " | ".join(reason_list)
            logger.warning(f"Vision validation failed: {reason}")
            return False, reason

        except json.JSONDecodeError:
            return False, f"JSON Decode Error from Vision Agent: {response_text[:100]}"