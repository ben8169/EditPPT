import json
from loguru import logger
from tools import TOOLS_SCHEMA, FUNCTION_MAP
from llm_client import call_llm
from prompts import *

class PPTAgent:
    def __init__(self, 
                prs,
                model: str):
        self.prs = prs
        self.model = model
        self.messages = [
            {
                "role": "system",
                "content": AGENT_SYSTEM_PROMPT
            }
        ]

    def run(self, plan_json: dict, parsed_json_list: list):
        # while True:
        for task in plan_json.get("tasks", []):
            page_number = task.get("page number")
            if not page_number:
                continue

            description = task.get("description", "")
            action = task.get("action", "")
            
            if parsed_json_list[int(page_number)-1].get("page number") == page_number:
                contents = parsed_json_list[int(page_number)-1].get("contents", [])
            else:
                contents = []

            self.messages.append({"role": "user", "content": create_agent_user_prompt(page_number, description, action, contents)})


            # success = False
            # retry_count = 0
            # max_retries = 1  # you can bump this up if you want more retries

            # while not success and retry_count < max_retries:
            #     attempt = retry_count + 1
            #     print(f"[Slide {page_number}] Attempt {attempt}/{max_retries} — sending prompt to LLM")
            #     prompt = create_agent_user_prompt(page_number, description, action, contents)

            response = call_llm(
                model=self.model,
                messages=self.messages,
                tools=TOOLS_SCHEMA,
                tool_choice="auto"
            )

            response_message = response.choices[0].message
            self.messages.append(response_message)

            if not response_message.tool_calls:
                return response_message.content

            for tool_call in response_message.tool_calls:
                function_name = tool_call.function.name
                function_args = json.loads(tool_call.function.arguments)
                
                logger.info(f"Tool Call: {function_name}({function_args})")
                result = self.execute_tool(function_name, function_args)
                
                self.messages.append({
                    "tool_call_id": tool_call.id,
                    "role": "tool",
                    "name": function_name,
                    "content": str(result),
                })

    def execute_tool(self, name, args):
        if name not in FUNCTION_MAP:
            return f"Error: Tool '{name}' not found."
        try:
            # Inject prs into the function call
            return FUNCTION_MAP[name](self.prs, **args)
        except Exception as e:
            logger.error(f"Execution Error: {e}")
            return f"Error: {str(e)}"
