import json
import time

from editppt.utils.llm_client import call_llm
from editppt.prompts import *
from editppt.utils.utils import parse_llm_response
from editppt.utils.logger_manual import TIMESTAMP, log_path


class Planner:
    def __init__(self, model, slide_name, total_slide_numbers):
        self.slide_name = slide_name
        self.total_slide_numbers = total_slide_numbers

        self.system_prompt = create_plan_prompt(
            self.slide_name,
            self.total_slide_numbers,
        )
        self.model = model

    def __call__(self, user_input: str):
        last_error_feedback = ""
        MAX_RETRIES = 3

        logs: list[str] = []
        had_error = False  

        def append_log(text: str):
            logs.append(text + "\n")

        for attempt in range(1, MAX_RETRIES + 1):
            try:
                call_llm_response = call_llm(
                    model=self.model,
                    messages=[
                        {"role": "system", "content": self.system_prompt},
                        {
                            "role": "user",
                            "content": last_error_feedback
                            + "\nNow, please create a plan for the following request:\n"
                            + user_input,
                        },
                    ],
                )

                response = call_llm_response.output_text

                append_log(
                    f"""[ATTEMPT {attempt}] RAW LLM RESPONSE
[USER INPUT]
{user_input}

[ERROR FEEDBACK]
{last_error_feedback}

[RAW LLM RESPONSE]
{response}
"""
                )

                plan, error = parse_llm_response(response)

                # 성공 + 이전에 에러 없었음 → 아무것도 생성하지 않음
                if error is None and not had_error:
                    return plan

                if error is None:
                    append_log(
                        f"""[ATTEMPT {attempt}] PARSED SUCCESS
{json.dumps(plan, indent=2, ensure_ascii=False)}
"""
                    )
                    break  

                # parsing error
                had_error = True
                e, state = error

                append_log(
                    f"""[ATTEMPT {attempt}] PARSE ERROR
[EXCEPTION]
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
                had_error = True

                append_log(
                    f"""[ATTEMPT {attempt}] EXCEPTION DURING LLM CALL
{type(e).__name__}: {e}
"""
                )

                print(f"[Attempt {attempt}/{MAX_RETRIES}] Exception during LLM call: {e}")
                time.sleep(1)

        if had_error:
            log_file_path = log_path(
                f"planner_{TIMESTAMP}.log",
                subdir="llm_debug_logs",
            )
            log_file_path.write_text("".join(logs), encoding="utf-8")

        if 'plan' in locals():
            return plan

        raise RuntimeError("Failed to obtain a valid plan from the LLM response.")
