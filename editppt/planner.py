from llm_client import call_llm
from prompts import PLAN_PROMPT
from utils import  parse_llm_response
import os
import json
from datetime import datetime
import time
import win32com.client

class Planner:
    def __init__(self, model):
        self.system_prompt = PLAN_PROMPT
        self.model = model
        # self.API_KEY = get_api_key_and_provider(model)[1]
    
    def __call__(self, user_input: str):
        # Construct the prompt for the LLM
        prompt = f"""
        {self.system_prompt}
        
        Now, please create a plan for the following request:
        {user_input}
        """
        # if "4.1" in model_name:
        #     response = _call_gpt_api(prompt= prompt, api_key=api_key, model = model_name)

        # response, input_tokens, output_tokens, total_cost = llm_request_with_retries_gemini(
        #     model_name = model_name,
        #     prompt_content = prompt
        # )
        #print(response)
        # The response should be a JSON string, but let's handle errors safely
        
        call_llm_response = call_llm(
            model=self.model,   
            messages=[{"role": "system", "content": prompt}]

        )
        response = call_llm_response.choices[0].message.content
        # input_tokens = call_llm_response.usage.input_tokens
        # output_tokens = call_llm_response.usage.output_tokens
        # total_cost = call_llm_response.usage.total_cost

        try:
            plan = parse_llm_response(response)
            return plan
            # return plan, input_tokens, output_tokens, total_cost
        except json.JSONDecodeError:
            # Fallback - return a basic structure with the raw response
            return {
                "understanding": "Failed to parse LLM response into proper JSON format",
                "tasks": [
                    {
                        "id": 1,
                        "description": "Review and manually interpret the plan",
                        "target": "N/A",
                        "action": "manual_review",
                        "details": response
                    }
                ],
                "requires_parsing": False,
                "requires_processing": False,
                "additional_notes": "LLM response was not in valid JSON format"
            }
        