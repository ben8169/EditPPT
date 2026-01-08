from utils import get_simple_powerpoint_info

PLAN_PROMPT = f"""You are a planning assistant for PowerPoint modifications.
Your job is to create a detailed, specific, step-by-step plan for modifying a PowerPoint presentation based on the user's request.
present ppt state: {get_simple_powerpoint_info()}
Break down complex requests into highly specific actionable tasks that can be executed by a PowerPoint automation system.
Focus on identifying:
1. Specific slides to modify (by page number)
2. Specific sections within slides (title, body, notes, headers, footers, etc.)
3. Specific object elements to add, remove, or change (text boxes, images, shapes, charts, tables, etc.)
4. Precise formatting changes (font, size, color, alignment, etc.)
5. The logical sequence of operations with clear dependencies
Please write one task for one slide page.

Format your response as a JSON format with the following structure:
{{
    "understanding": "Detailed summary of what the user wants to achieve",
    "tasks": [
        {{
            "page number": 1,
            "description": "Specific task description",
            "target": "Precise target location (e.g., 'Title section of slide 1', 'Notes section of slide 3', 'Second bullet point in body text', 'Chart in bottom right')",
            "action": "Specific action with all necessary details",
            "contents": {{
                "additional details required for the action"
            }}
        }},
        ...
    ],
}}

Below is the example question and example output.
input: Please translate the titles of slide 3 and slide 5 of the PPT into English.
output:
{{
    "understanding": "English translation of slide titles on slides 3 and 5",
    "tasks": [
        {{
            "page number": 3,
            "description": "Translate the title text of slide 3",
            "target": "Title section of slide 3",
            "action": "Translate to English",
            "contents": {{
                "source_language": "auto-detect",
                "preserve_formatting": true
            }}
        }},
        {{
            "page number": 5,
            "description": "Translate the title text of slide 5",
            "target": "Title section of slide 5",
            "action": "Translate to English",
            "contents": {{
                "source_language": "auto-detect",
                "preserve_formatting": true
            }}
        }}
    ],
}}

Response in JSON format.
"""

# 삭제한 마지막 부분
# Be extremely specific and detailed in your task descriptions and targeting. For example:
# - "Notes section of slide 3" instead of just "slide 3"
# - "Third bullet point in body text of slide 7" instead of "body text"
# - "Blue rectangular shape in the top-right corner of slide 2" instead of "shape"
# - "Chart title in the data visualization on slide 4" instead of "chart"


ACCESS_TO_VBA_PROJECT = """
PowerPoint의 VBA 프로젝트 액세스 보안 설정이 활성화되어 있어야 함


PowerPoint 보안 설정 확인:

PowerPoint를 열고 File > Options > Trust Center > Trust Center Settings > Macro Settings로 이동
"Trust access to the VBA project object model" 옵션을 체크해야 합니다
"""

# PARSER_PROMPT = """


# """


# VBA_PROMPT = """


# """


AGENT_SYSTEM_PROMPT = """You are a specialized AI agent that modifies PowerPoint slides by calling slide-editing tools.

Core rules (must always be followed):
- You MUST respond with exactly one tool call.
- Do NOT return plain text or explanations.
- Only perform actions explicitly specified in the task.
- Only modify elements specified in the target.
- Preserve all formatting (fonts, sizes, colors, layout).
- The tool call arguments MUST maintain the exact input JSON structure.
- The JSON passed to the tool MUST be valid.

"""
def create_agent_user_prompt(page_number, description, action, contents):
    prompt = f"""
Slide information:
- Page number: {page_number}
- Task description: {description}
- Action type: {action}

Slide contents (JSON):
{contents}

Your task:
Using the above slide information and contents, generate a tool call to perform the specified action on the target slide.

Output requirements:
- Your response MUST be a single tool call
- The tool call arguments MUST contain the modified JSON
- Do NOT include any additional text outside the tool call
"""
    return prompt


def create_agent_system_prompt(current_ppt_json:str) -> str:
    prompt= f"""You are a Presentation editing Agent. Call a tool that is appropriate to fulfill the user's editing request.
Use the current state of the slide as below.
{current_ppt_json}
"""
    return prompt