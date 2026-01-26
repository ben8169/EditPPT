import json
import time
import subprocess
import sys
import argparse
from pathlib import Path


from loguru import logger

import win32com.client

from editppt.utils.logger_manual import *
from editppt.utils.utils import *
from editppt.agent import *
from editppt.parser import Parser
from editppt.planner import Planner


logger = init_logger()

CURRENT_MODEL_NAME = "gpt-4.1"
CURRENT_VISION_MODEL_NAME = "gemini-2.5-pro"

class PPTContainer:
    def __init__(self, prs):
        self.prs = prs  

def kill_powerpoint_processes():
    try:
        subprocess.run(["taskkill", "/F", "/IM", "powerpnt.exe", "/T"], 
                       stdout=subprocess.DEVNULL, 
                       stderr=subprocess.DEVNULL)
        logger.info("기존의 모든 PowerPoint 프로세스를 정리했습니다.")
    except Exception as e:
        logger.warning(f"PowerPoint 프로세스 정리 중 오류 발생 (실행 중이 아닐 수 있음): {e}")


def initialize_ppt(ppt_path: Path):
    """
    Connects to PowerPoint and opens a specific file.
    """
    if not ppt_path.exists():
        raise FileNotFoundError(f"PPT file not found: {ppt_path}")

    try:
        ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception:
        logger.info("PowerPoint is not running. Launching new instance...")
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        ppt_app.Visible = True

    logger.info(f"Opening PPT file: {ppt_path}")
    prs = ppt_app.Presentations.Open(str(ppt_path))
    return prs

def parse_args():
    parser = argparse.ArgumentParser(
        description="PPT Editing Agent System"
    )
    parser.add_argument(
        "--file_path",
        type=str,
        required=True,
        help="Path to the PPTX file (absolute or relative)"
    )
    return parser.parse_args()

def main():
    args = parse_args()

    ppt_path = Path(args.file_path).expanduser().resolve()

    print("=== PPT Editing Agent System ===")
    logger.info(f"Resolved PPT path: {ppt_path}")

    kill_powerpoint_processes()
    time.sleep(1)

    try:
        prs = initialize_ppt(ppt_path)
        container = PPTContainer(prs)
        logger.info(
            f"Presentation [{container.prs.Name}] loaded "
            f"with {len(container.prs.Slides)} slides."
        )
    except Exception as e:
        logger.error(f"Failed to initialize PowerPoint: {e}")
        sys.exit(1)

    planner = Planner(
        model=CURRENT_MODEL_NAME,
        slide_name=container.prs.Name,
        total_slide_numbers=len(container.prs.Slides),
    )
    logger.info("Planner initialized")

    parser = Parser(
        container=container,
        total_slides=len(container.prs.Slides),
    )
    logger.info("Parser initialized")

    log_root = Path("logfiles") / TIMESTAMP
    log_root.mkdir(parents=True, exist_ok=True)

    (log_root / "parser_Database.json").write_text(
        json.dumps(parser.database, ensure_ascii=False, indent=4),
        encoding="utf-8",
    )

    edit_agent = EditAgent(
        container=container,
        model=CURRENT_MODEL_NAME,
    )

    vision_validator_agent = VisionValidatorAgent.create(
        activate_valid=True,
        container=container,
        model=CURRENT_MODEL_NAME,
    )
    logger.info("Agents initialized")

    print(f"\n[System] Agent is ready using model: {edit_agent.model}")

    # Agent loop
    while True:
        user_input = input("\n[User]: ").strip()
        if not user_input:
            continue

        if user_input == "eee":
            kill_powerpoint_processes()
            time.sleep(1)
            sys.exit(0)

        plan_json = planner(user_input)
        logger.info(f"Planner output received")

        (log_root / "planner.json").write_text(
            json.dumps(plan_json, ensure_ascii=False, indent=4),
            encoding="utf-8",
        )

        for task in plan_json.get("tasks", []):
            edit_agent.run(
                task=task,
                parser=parser,
                vision_validator_agent=vision_validator_agent,
            )


if __name__ == "__main__":
    main()