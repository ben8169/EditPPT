from llm_client import call_llm
from prompts import create_plan_prompt
from utils import  parse_llm_response
import os
import json
from datetime import datetime
import time
import win32com.client

from logtime import *

class Planner:
    def __init__(self, model, slide_name, total_slide_numbers):
        self.slide_name = slide_name
        self.total_slide_numbers = total_slide_numbers
        
        self.system_prompt = create_plan_prompt(self.slide_name, self.total_slide_numbers)
        self.model = model

    def __call__(self, user_input: str):
        last_error_feedback = ""
        MAX_RETRIES = 3

        for attempt in range(1, MAX_RETRIES + 1):


            try:
                # 1) Call LLM
                call_llm_response = call_llm(
                    model=self.model,
                    messages=[
                        {"role": "system", "content": self.system_prompt},
                        {"role": "user", "content": last_error_feedback +
                         "\nNow, please create a plan for the following request:\n" + user_input}
                    ]
                )

                response = call_llm_response.output_text

                # 🔹 로그: input + raw response
                dump_debug_log(
                    f"attempt_{attempt}_{TIMESTAMP}_raw.txt",
                    f"""[USER INPUT]
{user_input}

[ERROR FEEDBACK]
{last_error_feedback}

[RAW LLM RESPONSE]
{response}
"""
                )

                # 2) Parse response
                plan, error = parse_llm_response(response)

                # 3) Success
                if error is None:
                    dump_debug_log(
                        f"attempt_{attempt}_{TIMESTAMP}_parsed_success.txt",
                        json.dumps(plan, indent=2, ensure_ascii=False)
                    )
                    return plan

                # 4) Parsing failed
                e, state = error

                dump_debug_log(
                    f"attempt_{attempt}_parse_error.txt",
                    f"""[EXCEPTION]
{type(e).__name__}: {e}

[FAILED PAYLOAD / STATE]
{state}
"""
                )

                last_error_feedback = f"""
#### Additional Information for Correction ####
The previous response could not be parsed.

Error:
{type(e).__name__}: {e}

Invalid output:
{state}

Please fix the errors above and return ONLY valid JSON.
Do NOT include comments, explanations, or placeholders.
"""

                print(f"[Attempt {attempt}/{MAX_RETRIES}] Parsing failed. Retrying...")
                time.sleep(1)

            except Exception as e:
                # 🔹 LLM 호출 자체 실패 로그
                dump_debug_log(
                    f"attempt_{attempt}_exception.txt",
                    f"""[EXCEPTION DURING LLM CALL]
{type(e).__name__}: {e}
"""
                )

                print(f"[Attempt {attempt}/{MAX_RETRIES}] Exception during LLM call: {e}")
                time.sleep(1)

        # retries exhausted
        dump_debug_log(
            f"final_failure.txt",
            f"""[FINAL FAILURE]
Last error feedback:
{last_error_feedback}

Last response:
{response if 'response' in locals() else 'NO RESPONSE'}
"""
        )

        raise RuntimeError("Failed to obtain a valid plan from the LLM response.")

def dump_debug_log(filename: str, content: str):
    os.makedirs(f"logfiles/{TIMESTAMP}/llm_debug_logs", exist_ok=True)
    path = os.path.join(f"logfiles/{TIMESTAMP}/llm_debug_logs", filename)
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)