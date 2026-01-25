    # def _log_json(self, filename: str, data):
    #     path = os.path.join(self.log_dir, filename)
    #     with open(path, "w", encoding="utf-8") as f:
    #         json.dump(data, f, ensure_ascii=False, indent=4)

    # def _rollback_ppt(self, type: str, reason: str):
    #     """
    #     Close edited PPT and roll back to the last backup
    #     """
    #     logger.warning(f"{type} Feedback: {reason}")
    #     ppt_app = self.container.prs.Application
    #     self.container.prs.Close()
    #     time.sleep(0.5)
    #     self.container.prs = ppt_app.Presentations.Open(os.path.abspath(self.backup_path))

    # def run(self, task: dict, parser: object):
    #     page_number = task.get("page number")
    #     if not page_number:
    #         return

    #     tool_history = []
    #     feedback = []
    #     MAX_HISTORY = 4
    #     max_steps = 20
    #     step_count = 0

    #     # 초기 slide 상태 한 번만 읽기
    #     current_state = parser.process(page_number)

    #     while step_count < max_steps:
    #         step_count += 1

    #         # 1️⃣ LLM payload 구성 (항상 최신 current_state 사용)
    #         payload_message = [
    #             {"role": "system", "content": create_edit_agent_system_prompt(current_state)},
    #             {"role": "user", "content": create_edit_agent_user_prompt(
    #                 page_number,
    #                 task.get("description", ""),
    #                 task.get("action", ""),
    #                 task.get("contents", ""),
    #                 tool_history,
    #                 feedback,
    #                 MAX_HISTORY
    #             )}
    #         ]
    #         self.messages.append(payload_message)
    #         self._log_json(f"step_{step_count}_payload.json", payload_message)

    #         # 2️⃣ LLM 호출
    #         response = call_llm(
    #             model=self.model,
    #             messages=payload_message,
    #             tools=TOOLS_SCHEMA,
    #             tool_choice="auto"
    #         )

    #         # 3️⃣ 단일 tool 실행
    #         tool_executed = False
    #         for item in response.output:
    #             if item.type == "function_call":
    #                 function_args = json.loads(item.arguments) if isinstance(item.arguments, str) else item.arguments
    #                 function_name = item.name

    #                 # Undo 처리
    #                 if function_name == "undo_action":
    #                     reason = function_args.get("reason", "User requested undo")
    #                     self._execute_tool("undo_action", {"reason": reason, "container": self.container})
    #                     # rollback 후 slide 상태 갱신
    #                     current_state = parser.process(page_number)
    #                     self._log_json(f"step_{step_count}_undo.json", {"reason": reason, "slide_state": current_state})
    #                     tool_executed = True
    #                     break

    #                 # tool 실행 전 slide_json 포함
    #                 logger.info(f"Tool Call: {function_name}({function_args})")

    #                 if function_name in ["set_text_style_preserve_runs", "replace_text_from_textbox"]:
    #                     function_args["slide_json"] = current_state
    #                     if function_name == "set_text_style_preserve_runs":
    #                         for key in ["bold", "italic", "underline", "font_name", "font_size"]:
    #                             if key in function_args and function_args[key] is False:
    #                                 del function_args[key]

    #                 # tool 실행
    #                 self._execute_tool(function_name, function_args)
    #                 tool_executed = True

    #                 # 실행 후 상태 갱신 (한 번만)
    #                 current_state = parser.process(page_number)

    #                 # tool_history 업데이트
    #                 tool_history.append({
    #                     "tool_name": function_name,
    #                     "arguments": function_args,
    #                     "result_state": current_state,
    #                     "call_id": getattr(item, "call_id", None)
    #                 })
    #                 if len(tool_history) > MAX_HISTORY:
    #                     tool_history.pop(0)
    #                 self._log_json(f"step_{step_count}_tool_execution.json", tool_history[-1])

    #                 break  # 한 번에 하나만 실행

    #         self._log_json(f"step_{step_count}_feedback.json", feedback)            
    #         # 4️⃣ 종료 조건
    #         if not tool_executed:
    #             print(f"Step {step_count}: LLM indicated no further tools needed. Stopping.")
    #             break
    #         self._log_json("slide_final_state.json", current_state)
    #         self._log_json(f"agent_Message.json", self.messages)
    #         print(f"All logs saved in {self.log_dir}")



#Tools.py
    "undo_action":undo_action,

 {
        "type": "function",
        "name": "undo_action",
        "description": "Undo the last tool action and rollback the slide to the previous state.",
        "parameters": {
            "type": "object",
            "properties": {
                "reason": {"type": "string", "description": "Reason for undo, e.g., incorrect result"}
            },
            "required": ["reason"]
        }
    },




# insert, delete 없애보자
  {
    "type": "function",
    "name": "insert_text_from_textbox",
    "description": "Insert new text into a text container while inheriting the surrounding text style.",
    "parameters": {
        "type": "object",
        "properties": {
        "slide_number": {
            "type": "integer",
            "description": "1-based slide index"
        },
        "shape_id": {
            "type": "integer",
            "description": "PowerPoint Shape ID"
        },
        "preceding_text": {
            "type": "string",
            "description": "Text that appears immediately before the insertion point (used as an anchor)"
        },
        "char_start_index": {
            "type": "integer",
            "description": "0-based character index near the insertion position"
        },
        "new_text": {
            "type": "string",
            "description": "Text to insert"
        },
        "container": {
            "type": "string",
            "enum": ["shape", "table_cell"],
            "default": "shape"
        },
        "row_index": {
            "type": "integer",
            "description": "1-based row index (required if container=table_cell)"
        },
        "col_index": {
            "type": "integer",
            "description": "1-based column index (required if container=table_cell)"
        }
        },
        "required": [
        "slide_number",
        "shape_id",
        "preceding_text",
        "char_start_index",
        "new_text"
        ]
    }
    },
    {
    "type": "function",
    "name": "delete_text_from_textbox",
    "description": "Delete a specific text segment from a text container without affecting surrounding styles.",
    "parameters": {
        "type": "object",
        "properties": {
        "slide_number": {
            "type": "integer",
            "description": "1-based slide index"
        },
        "shape_id": {
            "type": "integer",
            "description": "PowerPoint Shape ID"
        },
        "target_text": {
            "type": "string",
            "description": "Exact text to delete"
        },
        "char_start_index": {
            "type": "integer",
            "description": "0-based character index where the text is expected to start"
        },
        "container": {
            "type": "string",
            "enum": ["shape", "table_cell"],
            "default": "shape"
        },
        "row_index": {
            "type": "integer",
            "description": "1-based row index (required if container=table_cell)"
        },
        "col_index": {
            "type": "integer",
            "description": "1-based column index (required if container=table_cell)"
        }
        },
        "required": [
        "slide_number",
        "shape_id",
        "target_text",
        "char_start_index"
        ]
    }
    },


    {
    "type": "function",
    "name": "replace_text_from_textbox",
    "description": "Replace a specific text segment while preserving its original font style.",
    "parameters": {
        "type": "object",
        "properties": {
        "slide_number": {
            "type": "integer",
            "description": "1-based slide index"
        },
        "shape_id": {
            "type": "integer",
            "description": "PowerPoint Shape ID"
        },
        "target_text": {
            "type": "string",
            "description": "Exact text to be replaced"
        },
        "char_start_index": {
            "type": "integer",
            "description": "0-based character index where the text is expected to start"
        },
        "new_text": {
            "type": "string",
            "description": "Replacement text"
        },
        "container": {
            "type": "string",
            "enum": ["shape", "table_cell"],
            "default": "shape"
        },
        "row_index": {
            "type": "integer",
            "description": "1-based row index (required if container=table_cell)"
        },
        "col_index": {
            "type": "integer",
            "description": "1-based column index (required if container=table_cell)"
        }
        },
        "required": [
        "slide_number",
        "shape_id",
        "target_text",
        "char_start_index",
        "new_text"
        ]
    }
    },