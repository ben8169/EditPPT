import os
import json
import time
import win32com.client
from loguru import logger
import subprocess

from logtime import *
from agent import *
from parser import Parser
from planner import Planner
from utils import *

import sys


# PPT_PATH = "../sample_ppt/kairi.pptx"  
# PPT_PATH = "../sample_ppt/news_card.pptx"  
PPT_PATH = "../sample_ppt/math_simple.pptx"  
CURRENT_MODEL_NAME = "gpt-4.1"
CURRENT_VISION_MODEL_NAME = "gemini-2.5-pro"

class PPTContainer:
    def __init__(self, prs):
        self.prs = prs  

def kill_powerpoint_processes():
    try:
        # /F: 강제 종료, /IM: 이미지 이름 지정, /T: 자식 프로세스까지 종료
        # > nul 2>&1 은 실행 결과 메시지를 콘솔에 출력하지 않게 합니다 (Windows 전용)
        subprocess.run(["taskkill", "/F", "/IM", "powerpnt.exe", "/T"], 
                       stdout=subprocess.DEVNULL, 
                       stderr=subprocess.DEVNULL)
        logger.info("기존의 모든 PowerPoint 프로세스를 정리했습니다.")
    except Exception as e:
        logger.warning(f"PowerPoint 프로세스 정리 중 오류 발생 (실행 중이 아닐 수 있음): {e}")



def initialize_ppt(file_path=None):
    """
    Connects to PowerPoint and opens a specific file.
    """
    try:
        # Connect to existing PowerPoint instance
        ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
    except:
        # Open a new PowerPoint instance
        logger.info("PowerPoint is not running. Launching new instance...")
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        ppt_app.Visible = True

    if file_path and os.path.exists(file_path):
        # Open the specific file provided
        abs_path = os.path.abspath(file_path)
        logger.info(f"Opening file: {abs_path}")
        prs = ppt_app.Presentations.Open(abs_path)
    else:
        # Create a new presentation if no path is provided or file doesn't exist
        logger.info("No valid file path provided.")
        raise   
    
    return prs

def main():
    print("=== PPT Editing Agent System ===")
    kill_powerpoint_processes()
    time.sleep(1)

    try:
        prs = initialize_ppt(PPT_PATH)
        container = PPTContainer(prs)
        logger.info(f"Presentation name [{container.prs.Name}] - loaded with {len(container.prs.Slides)} slides.")
    except Exception as e:
        logger.error(f"Failed to initialize PowerPoint: {e}")
        return

    planner = Planner(model=CURRENT_MODEL_NAME)
    logger.info(f"Planner 완료 ...")
    
    parser = Parser(container=container, total_slides=len(container.prs.Slides)) 
    logger.info(f"Parser 완료 ...")

    os.makedirs(f"logfiles/{TIMESTAMP}", exist_ok=True)

    with open(f"logfiles/{TIMESTAMP}/parser_Database.json", "w", encoding="utf-8") as f:
        json.dump(parser.database, f, ensure_ascii=False, indent=4)

    edit_agent = EditAgent(container=container, model=CURRENT_MODEL_NAME)
    
    vision_validator_agent = VisionValidatorAgent.create(
        activate_valid=True, 
        container=container, 
        model=CURRENT_MODEL_NAME
    )
    logger.info(f"Agent 완료 ...")

    print(f"\n[System] Agent is ready using model: {edit_agent.model}")

    # 3. Agent Loop
    while True:
        user_input = input("\n[User]: ")
        if not user_input.strip():
            continue

        if user_input.strip() == 'eee':
            try:
                kill_powerpoint_processes()
                time.sleep(1)
            except Exception as e:
                print(f"PPT 닫는 중 에러 발생: {e}")
            sys.exit(0)

        plan_json = planner(user_input)
        logger.info(f"Planner 완료 - 결과: {str(plan_json)[:100]}...")
        with open(f"logfiles/{TIMESTAMP}/planner.json", "w", encoding="utf-8") as f:
            json.dump(plan_json, f, ensure_ascii=False, indent=4)

        for task in plan_json.get("tasks", []):
            edit_agent.run(
                task=task,
                parser=parser,
                vision_validator_agent=vision_validator_agent
            )

if __name__ == "__main__":
    main()