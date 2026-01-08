import os
import json
import win32com.client
from loguru import logger
from agent import PPTAgent
# from parser import parse_ppt
from parser import Parser
from planner import Planner
from utils import *


PPT_PATH = "../sample_ppt/math_simple.pptx"  
# CURRENT_MODEL_NAME = "solar-pro2"
CURRENT_MODEL_NAME = "gpt-4.1"

def initialize_ppt(file_path=None):
    """
    Connects to PowerPoint and opens a specific file or creates a new one.
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
        logger.info("No valid file path provided. Creating a new presentation...")
        prs = ppt_app.Presentations.Add()
        prs.Slides.Add(1, 12)  # 12 = ppLayoutBlank
    
    return prs

# def get_presentation_info(prs):
#     """Returns the full structure of the current PPT as a JSON string."""
#     data = parse_ppt(prs)
#     return data


def main():
    print("=== PPT Editing Agent System ===")
        
    try:
        prs = initialize_ppt(PPT_PATH)
        # prs_dict = Parser(baseline=True).process()
        # print(prs_dict)
        logger.info(f"Presentation loaded with {len(prs.Slides)} slides.")
    except Exception as e:
        logger.error(f"Failed to initialize PowerPoint: {e}")
        return

    # 2. Create planer and agent
    planner = Planner(model=CURRENT_MODEL_NAME)
    parser = Parser(baseline=True)
    parsed_json_list = parser.process()
    logger.info(f"Parser 완료 ...")
    with open("parser_json.json", "w", encoding="utf-8") as f:
        for pj in parsed_json_list:
            json.dump(pj, f, ensure_ascii=False, indent=4)

    agent = PPTAgent(prs=prs, model=CURRENT_MODEL_NAME)

    print(f"\n[System] Agent is ready using model: {agent.model}")

    # 3. Chat Loop
    while True:
        user_input = input("\n[User]: ")
        plan_json = planner(user_input)
        logger.info(f"Planner 완료 - 결과: {str(plan_json)[:100]}...")
        with open("planner_json.json", "w", encoding="utf-8") as f:
            json.dump(plan_json, f, ensure_ascii=False, indent=4)

        agent.run(plan_json, parsed_json_list)
        if not user_input.strip():
            continue

if __name__ == "__main__":
    main()
