import win32com.client
import os
import base64
from dotenv import load_dotenv
from google import genai 
from google.genai import types
from openai import OpenAI
import json

load_dotenv()

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")

QA_SCHEMA = {
  "type": "object",
  "properties": {
    "HasCriticalIssues": {
      "type": "string",
      "enum": ["Yes", "No"]
    },
    "Issues": {
      "type": "array",
      "items": {
        "type": "object",
        "properties": {
          "IssueType": {
            "type": "string",
            "enum": [
              "TEXT_OVERFLOW",
              "ELEMENT_COLLISION",
              "ALIGNMENT_INCONSISTENCY",
              "LEGIBILITY_CONTRAST"
            ]
          },
          "AffectedShapeIDs": {
            "type": "array",
            "items": { "type": "string" }
          },
          "TechnicalDiagnosis": {
            "type": "string"
          },
          "ActionableFix": {
            "type": "string"
          }
        },
        "required": [
          "IssueType",
          "AffectedShapeIDs",
          "TechnicalDiagnosis",
          "ActionableFix"
        ],
        "additionalProperties": False
      }
    }
  },
  "required": ["HasCriticalIssues", "Issues"],
  "additionalProperties": False
}


import json
from json import JSONDecodeError


def load_possible_multiple_json(path: str):
    """Load a JSON file that may contain either a single JSON value
    or several JSON objects concatenated together (no surrounding array).
    Returns either a single Python object or a list of objects.
    """
    with open(path, 'r', encoding='utf-8') as f:
        s = f.read().strip()

    try:
        # Normal case: the file is valid JSON (object or array)
        return json.loads(s)
    except JSONDecodeError:
        # Try to fix concatenated objects like: ...}{... by inserting commas
        fixed = s.replace('}\n{', '},\n{').replace('}{', '},\n{')
        try:
            return json.loads('[' + fixed + ']')
        except JSONDecodeError as e:
            # Re-raise with more context for easier debugging
            raise JSONDecodeError(f"Failed to parse JSON even after fixing concatenation: {e}", s, e.pos)





def export_all_slides_to_images(output_folder="SlideScreenshots"):
    try:
        # 1. 실행 중인 파워포인트 앱 연결
        ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        presentation = ppt_app.ActivePresentation
        
        # 절대 경로로 변환하여 혼선 방지
        abs_output_path = os.path.abspath(output_folder)
        
        # 2. 저장할 폴더 생성 (수정: s.path -> os.path)
        if not os.path.exists(abs_output_path):
            os.makedirs(abs_output_path)
            print(f"Folder created: {abs_output_path}")

        # 3. 전체 슬라이드 개수 파악
        slide_count = presentation.Slides.Count
        print(f"Total slides to export: {slide_count}")

        # 4. 루프를 돌며 각 슬라이드 저장
        for i in range(1, slide_count + 1):
            slide = presentation.Slides(i)
            file_name = f"slide_{i}.png"
            file_path = os.path.join(abs_output_path, file_name)
            
            # Export(경로, 포맷)
            slide.Export(file_path, "PNG")
            print(f"[{i}/{slide_count}] Exported: {file_name}")

        print("\nAll slides have been exported successfully.")
        return abs_output_path
        
    except Exception as e:
        print(f"Error during export: {e}")
        return None

def encode_image(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')

# --- 2. 모델별 호출 함수 정의 ---

def check_design_gemini(image_path, prompt, api_key):
    try:
        if not os.path.exists(image_path):
            return f"[Gemini Error] File not found: {image_path}"
            
        client = genai.Client(api_key=api_key)
        
        with open(os.path.abspath(image_path), "rb") as f:
            img_bytes = f.read()

        # 모델명 확인 필요 (최신 모델명으로 교체)
        response = client.models.generate_content(
            model="gemini-2.5-pro", 
            contents=[
                types.Part.from_bytes(data=img_bytes, mime_type="image/png"),
                prompt
            ]
        )
        return f"[Gemini Feedback]\n{response.text}"
    except Exception as e:
        return f"[Gemini Error] {str(e)}"

def check_design_gpt(image_path, prompt, api_key):
    try:
        if not os.path.exists(image_path):
            return f"[GPT Error] File not found: {image_path}"
            
        client = OpenAI(api_key=api_key)
        base64_image = encode_image(image_path)
        
        response = client.responses.create(
    model="gpt-5",
    input=[
        {
            "role": "user",
            "content": [
                {"type": "input_text", "text": prompt},
                {
                    "type": "input_image",
                    "image_url": f"data:image/png;base64,{base64_image}"
                }
            ]
        }
    ],
    text={
        "format": {
            "name": "ppt_qa_result",
            "type": "json_schema",
            "schema": QA_SCHEMA
        }
    }
)
        return f"[GPT Feedback]\n{response.output_text}"
    except Exception as e:
        return f"[GPT Error] {str(e)}"


# --- 3. 실행 영역 ---

if __name__ == "__main__":
    # 1. 먼저 슬라이드를 이미지로 추출합니다.
    # output_dir = "SlideScreenshots"
    # exported_path = export_all_slides_to_images(output_dir)
    with open('parser_json.json', 'r', encoding='utf-8') as f:
        parsed_slide = load_possible_multiple_json(r'C:\Users\ben81\Documents\EditPPT\editppt\logfiles\20260120_023547\parser_Database.json')

    # for i in range(5,len(os.listdir('.SlideScreenshots'))+1):    
    k = 6 
    for i in range(k, k+1):     
    # for i in range(7, 8):     

        # 추출된 파일 중 7번째 슬라이드 경로 설정
        IMAGE_PATH = os.path.join('.SlideScreenshots', f"slide_{i}.png")

        parsed_contents = parsed_slide.get(f'{i}')
        agent_request = f'Translate all English into Korean in slide {i}'
        
        PROMPT = f"""You are a professional PPT Quality Assurance (QA) specialist.

Your role is NOT to review the entire slide.
Your role is to validate ONLY the visual and structural impact
caused by a specific modification request executed by another agent.

You MUST strictly limit your evaluation scope to:
- The shapes explicitly modified by the agent request
- Shapes spatially adjacent to or directly affected by those modifications

You MUST NOT evaluate or report issues unrelated to the modification,
even if they would normally qualify as critical issues.

---

### Agent Modification Context (SOURCE OF SCOPE)

The following request describes WHAT was modified and WHERE:

{agent_request}

This request DEFINES your inspection boundary.
Anything outside the logical or spatial impact of this request
MUST be ignored.

---

### Evaluation Objective (CHANGE IMPACT ONLY)

Your task is to determine whether the requested modification
introduced any NEW, SEVERE presentation defects in the affected area.

You are ONLY checking for regressions caused by the change, such as:
- Newly introduced text overflow
- Newly introduced element overlap
- Newly introduced severe misalignment
- Newly introduced legibility issues

Pre-existing issues that were already present before the modification
MUST NOT be reported.

If an issue is subtle, borderline, subjective, or unrelated to the change,
you MUST ignore it.

If there is any uncertainty about whether an issue was caused by the change,
you MUST assume it was NOT caused by the change.

---

### Criticality Threshold (HARD SUPPRESSION RULES)

You MUST NOT report an issue unless ALL of the following are true:

1. The issue is directly caused by the modification described in 'agent_request'.
2. The issue is immediately obvious to a human viewer at normal presentation scale.
3. The issue clearly degrades readability or professional credibility.

DO NOT report issues that:
- Exist outside the modified area
- Existed prior to the modification
- Require close inspection, measurement, or interpretation
- Could reasonably be intentional design choices

If in doubt, report NO issues.

---

### Source of Truth Rules (Image vs JSON)

1. **Rendered Visual Outcome (Primary)**
   - If the affected shapes appear visibly broken in the image,
     the image is the primary source of truth.

2. **Structural / Z-Order Issues (Secondary)**
   - If the agent request modified structure (size, z-order, visibility),
     JSON may be used to confirm regressions even if subtle in the image.

If the defect cannot be clearly linked to the modification,
DO NOT report it.

---

### Allowed Issue Types (ENUM – MUST USE EXACTLY)

- TEXT_OVERFLOW
- ELEMENT_COLLISION
- ALIGNMENT_INCONSISTENCY

---

### Output Requirements (STRICT)

- **HasCriticalIssues** must be `"Yes"` or `"No"`.

#### If **HasCriticalIssues = "Yes"**
Return ONLY issues that:
- Were introduced by the modification
- Affect the modified or directly adjacent shapes

Each issue MUST include:
- **IssueType**
- **AffectedShapeIDs**
- **TechnicalDiagnosis** (explain how the modification caused it)
- **ActionableFix** (how to fix the regression)

#### If **HasCriticalIssues = "No"**
Return ONLY:
{{
  "HasCriticalIssues": "No"
}}

Do NOT return any additional commentary.

---

### Fail-Safe Rule (MANDATORY)

If you cannot confidently attribute a defect to the modification,
you MUST return:

{{
  "HasCriticalIssues": "No"
}}

---
### Agent request
{agent_request}
### Slide Contents JSON 
{parsed_contents}

"""
        # PROMPT += json.dumps(parsed_contents, indent=2, ensure_ascii=False)
        # 이미지 파일이 실제로 존재하는지 확인 후 실행
        if os.path.exists(IMAGE_PATH):
            print(check_design_gemini(IMAGE_PATH, PROMPT, GEMINI_API_KEY))
            print("-" * 30)
            print(check_design_gpt(IMAGE_PATH, PROMPT, OPENAI_API_KEY))
            print("-" * 30)
            print()
        else:
            print(f"오류: '{IMAGE_PATH}' 파일을 찾을 수 없습니다. 슬라이드 개수를 확인하세요.")