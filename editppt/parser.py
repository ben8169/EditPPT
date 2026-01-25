from utils import parse_active_slide_objects
import os
import json
from datetime import datetime
import time
import win32com.client
import logging
import copy
from copy import deepcopy
from logtime import *
import re

from llm_client import call_llm
from prompts import *


class Parser:
    def __init__(self, container: object, total_slides: int):
        """
        Args:
            container (object): PPTContainer instance containing the prs.
            total_slides (int): Total number of slides.
        """
        self.database = {}
        self.edit_history = {}
        self.container = container  
        self.total_slides = total_slides

        # for page_num in range(1, min(10, total_slides + 1)):
        #     self.database[page_num] = parse_active_slide_objects(page_num, self.container.prs)
        #     print(f"Parsed slide {page_num}/{total_slides}")


    def process(self, page_number: int):
        # if page_number < 1 or page_number > self.total_slides:
        #     raise ValueError("Invalid page number")
        
        # if page_number not in self.database:
        print(f'Parsing Page {page_number}...')
        print('='*40)
        self.database[page_number] = parse_active_slide_objects(page_number, self.container.prs)
        with open(f"logfiles/{TIMESTAMP}/parser_Database.json", "w", encoding="utf-8") as f:
            json.dump(self.database, f, ensure_ascii=False, indent=4)

        return self.database[page_number]

    def update_after_edit(self,
                        text_validation: bool,
                        model: str,
                        page_number: int,
                        description: str,
                        action: str,
                        detailed_contents: str,
                        used_tools: list) -> bool:
        """
        PPT 수정 결과를 LLM이 검증
        True = 요청사항 만족, False = undo 
        """

        old_parse = self.database.get(page_number, None)
        if old_parse is None:
            raise RuntimeError("Slide not parsed by parser.process()")
        new_parse = parse_active_slide_objects(page_number, self.container.prs)

        # 수정 전 데이터 기록 (Append 모드)
        with open(f"logfiles/{TIMESTAMP}/oldparse_{page_number}.txt", "a", encoding="utf-8") as f:
            f.write(f"\n{'='*50}\n")
            f.write(json.dumps(old_parse, ensure_ascii=False, indent=4))
            f.write("\n")

        # 수정 후 데이터 기록 (Append 모드)
        with open(f"logfiles/{TIMESTAMP}/newparse_{page_number}.txt", "a", encoding="utf-8") as f:
            f.write(f"\n{'='*50}\n")
            f.write(json.dumps(new_parse, ensure_ascii=False, indent=4))
            f.write("\n")

        #validation 1
        if new_parse == old_parse:
            if used_tools:
                return False, "Tool used, but no differences. Check the target slide number or shape id carefully."
            else:
                return False, "Unknown Error."
        else:
            if text_validation:
                # LLM 검증
                messages = [
                    {"role": "system",
                    "content": create_text_validator_agent_system_prompt(
                        page_number, description, action, detailed_contents)},
                    {"role": "user",
                    "content": create_text_validator_agent_user_prompt(new_parse, old_parse, used_tools)}
                ]
                response = call_llm(model=model, messages=messages)
                response_text = (response.output_text or "").strip()

                if response_text.lower().startswith("true"):
                    valid = True
                    reason = None
                else:
                    valid = False
                    if "|" in response_text:
                        reason = response_text.split("|")[1].strip()
                    else:
                        reason = response_text.replace("False", "").strip(": ").strip()

                if valid:
                    return True, None
                else:
                    reason = re.sub(r"^(false|no)[:.\s]*", "", response_text, flags=re.IGNORECASE).strip()
                    return False, reason
            else:
                self.edit_history.setdefault(page_number, []).append(deepcopy(old_parse))
                self.database[page_number] = new_parse
                return True, None
