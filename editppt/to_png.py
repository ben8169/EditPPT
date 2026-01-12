import win32com.client
import os
import base64
from dotenv import load_dotenv
from google import genai 
from google.genai import types
from openai import OpenAI
from anthropic import Anthropic  # 추가: Anthropic 라이브러리 임포트
import json

load_dotenv()

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")



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
        
        response = client.chat.completions.create(
            model="gpt-5",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {
                            "type": "image_url", 
                            "image_url": {"url": f"data:image/png;base64,{base64_image}"}
                        }
                    ]
                }
            ]
        )
        return f"[GPT-4o Feedback]\n{response.choices[0].message.content}"
    except Exception as e:
        return f"[GPT-4o Error] {str(e)}"

def check_design_claude(image_path, prompt, api_key):
    """
    Send an image to Claude for analysis.
    
    Args:
        image_path: Path to the image file
        prompt: Text prompt describing what you want Claude to analyze
        api_key: Your Anthropic API key
        
    Returns:
        String containing Claude's feedback or error message
    """
    try:
        if not os.path.exists(image_path):
            return f"[Claude Error] File not found: {image_path}"
        
        # Determine media type from file extension
        ext = os.path.splitext(image_path)[1].lower()
        media_type_map = {
            '.png': 'image/png',
            '.jpg': 'image/jpeg',
            '.jpeg': 'image/jpeg',
            '.gif': 'image/gif',
            '.webp': 'image/webp'
        }
        media_type = media_type_map.get(ext, 'image/png')
        
        client = Anthropic(api_key=api_key)
        base64_image = encode_image(image_path)
        
        response = client.messages.create(
            model="claude-sonnet-4-5-20250929",  # Updated to latest model
            max_tokens=1000,
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": media_type,
                            "data": base64_image
                        }
                    },
                    {"type": "text", "text": prompt}
                ]
            }]
        )
        return f"[Claude Feedback]\n{response.content[0].text}"
    except Exception as e:
        return f"[Claude Error] {str(e)}"

# --- 3. 실행 영역 ---

if __name__ == "__main__":
    # 1. 먼저 슬라이드를 이미지로 추출합니다.
    output_dir = "SlideScreenshots"
    exported_path = export_all_slides_to_images(output_dir)
    with open('parser_json.json', 'r', encoding='utf-8') as f:
        parsed_slide = load_possible_multiple_json('parser_json.json')

    for i in range(1,len(os.listdir('SlideScreenshots'))+1):     
        if i != 7:
            continue   
        # 추출된 파일 중 7번째 슬라이드 경로 설정
        IMAGE_PATH = os.path.join('SlideScreenshots', f"slide_{i}.png")

        parsed_contents = parsed_slide[i-1].get('contents')
        
        PROMPT = f"""You are a professional PPT Quality Assurance (QA) specialist. Your primary objective is to detect technical and structural defects in the slide that compromise professional standards and readability. 

Perform a rigorous audit by cross-referencing the slide image with the provided JSON data. Focus exclusively on identifying "Critical Violations" in the following categories:

1. **Text Overflow & Clipping**: 
   - Detect text exceeding its bounding box or the slide margins.
   - Flag unintended line breaks, truncated text, or hidden characters caused by excessive font size or insufficient box dimensions.

2. **Element Overlap & Collision**: 
   - Identify unintended intersections between text, shapes, or icons that obscure information.
   - Verify that the stacking order (z-index) is correct (e.g., text must never be obscured by background shapes).

3. **Alignment & Spacing Inconsistency**: 
   - Check if elements intended to be aligned (e.g., headers, bullet points) have inconsistent coordinates.
   - Detect uneven gutters or margins between repetitive components (e.g., icon sets or grid layouts).

4. **Legibility & Contrast**: 
   - Flag text that is visually inaccessible due to poor color contrast against the background or placement over busy image areas.

**Output Requirements**:
For every issue detected, you must return a structured response in the following format:
- **Issue Type**: (e.g., Text Overflow, Element Collision, etc.)
- **Affected Element ID(s)**: (Reference the specific ID(s) from the provided JSON)
- **Technical Diagnosis**: A precise explanation of the structural error (e.g., "Text is clipped at the bottom by 12px due to small bounding box").
- **Actionable Fix**: A specific technical instruction for correction (e.g., "Reduce font size to 14pt" or "Adjust X-coordinate to 240").

Slide Contents JSON: 
{parsed_contents}
"""
        # 이미지 파일이 실제로 존재하는지 확인 후 실행
        if os.path.exists(IMAGE_PATH):
            print(check_design_gemini(IMAGE_PATH, PROMPT, GEMINI_API_KEY))
            print("-" * 30)
            print(check_design_gpt(IMAGE_PATH, PROMPT, OPENAI_API_KEY))
            print("-" * 30)
            print(check_design_claude(IMAGE_PATH, PROMPT, ANTHROPIC_API_KEY))
        else:
            print(f"오류: '{IMAGE_PATH}' 파일을 찾을 수 없습니다. 슬라이드 개수를 확인하세요.")