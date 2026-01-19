import win32com.client
import pywintypes
import openai
from openai import OpenAI

import re
import json
import ast
import time

from logtime import *
from msoffice_map import *


def parse_llm_response(response):
    """
    Parse JSON or Python-like structures from an LLM response.

    Returns:
    - (parsed_obj, None) on success
    - (None, (exception, payload_or_response)) on failure
    """

    # 1. 입력 검증
    if not response or not isinstance(response, str):
        e = ValueError("response is empty or not a string")
        return None, (e, response)

    # 2. 마크다운 코드펜스 제거
    response_clean = re.sub(r'```(?:json)?', '', response).strip()

    # 3. JSON / list 추출
    match = re.search(r'(\{.*\}|\[.*\])', response_clean, re.DOTALL)
    if not match:
        e = ValueError("No JSON object could be decoded")
        return None, (e, response_clean)

    payload = match.group(1)

    # 4. trailing comma 제거
    payload = re.sub(r',\s*([\}\]])', r'\1', payload)

    try:
        parsed = json.loads(payload)
        return parsed, None
    except json.JSONDecodeError as e_json:
        # 7. Python literal fallback
        try:
            parsed = ast.literal_eval(payload)
            return parsed, None
        except Exception as e_ast:
            return None, (e_ast, payload)




def extract_content_after_edit(plan_json):
    result = []
    
    if 'tasks' in plan_json and len(plan_json['tasks']) > 0:
        for task in plan_json['tasks']:
            if 'content after edit' in task and isinstance(task['content after edit'], list):
                result.extend(task['content after edit'])
    
    return result

def extract_last_text_content(plan_json):
    last_text = ""
    
    if 'tasks' in plan_json and len(plan_json['tasks']) > 0:
        for task in plan_json['tasks']:
            if 'contents' in task:
                contents_str = task['contents']
                # Text content: 패턴을 모두 찾아서 리스트로 만듦
                text_contents = re.findall(r'Text content: (.*?)(?=\n\s+Font:|$)', contents_str, re.DOTALL)
                
                # 마지막 Text content: 내용을 반환 (없으면 빈 문자열)
                if text_contents:
                    last_text = text_contents[-1].strip()
    
    return last_text

def create_thinking_queue(plan_json):
    # thinking queue
    temp_tasks = []
    temp_actions = []
    
    print_data_ = ""

    for i in range(len(plan_json['tasks'])):
        temp_tasks.append(plan_json['tasks'][i]['target'])
        temp_actions.append(plan_json['tasks'][i]['action'])
    
    for i in range(len(temp_tasks)):
        print_data_ += f"• {temp_actions[i]} 작업을 '{temp_tasks[i]}'에 적용합니다.\n"
    
    return print_data_


import openai
from openai import OpenAI
import tiktoken

# 모델별 토큰당 단가(예시: USD/1K tokens)
PRICING = {
    #"gpt-4.1-2025-04-14":    {"prompt": 0.03/1000, "completion": 0.06/1000},
    "gpt-4.1-mini-2025-04-14":{"prompt": 0.4/1000000, "completion": 1.6/1000000},
    #"gpt-4.1-nano-2025-04-14":{"prompt": 0.001/1000, "completion": 0.001/1000},
    #"o4-mini":               {"prompt": 0.002/1000, "completion": 0.002/1000},
}

def count_tokens(text: str, model: str) -> int:
    """tiktoken으로 토큰 수 계산"""
    try:
        enc = tiktoken.encoding_for_model(model)
    except KeyError:
        enc = tiktoken.get_encoding("cl100k_base")
    return len(enc.encode(text))

def _call_gpt_api(prompt: str, api_key: str, model: str):
    # --- API 키 설정 및 모델 검증/매핑 ---
    openai.api_key = api_key

    allowed = ["gpt-4.1", "gpt-4.1-mini", "gpt-4.1-nano", "o4-mini"]
    if model not in allowed:
        raise ValueError(f"Model must be one of {allowed}")

    if model == "gpt-4.1":
        model = "gpt-4.1-2025-04-14"
    elif model == "gpt-4.1-mini":
        model = "gpt-4.1-mini-2025-04-14"
    elif model == "gpt-4.1-nano":
        model = "gpt-4.1-nano-2025-04-14"
    # o4-mini는 그대로

    # --- API 호출 ---
    client = OpenAI(api_key=api_key)
    response = client.responses.create(
        model=model,
        instructions="You are a coding assistant that edits PowerPoint slides.",
        input=prompt,
    )
    text = response.output_text

    # --- 토큰 수 계산 (usage 필드가 있으면 그걸 쓰고, 없으면 count_tokens) ---
    if getattr(response, "usage", None):
        inp_toks = response.usage.input_tokens
        out_toks = response.usage.output_tokens
    else:
        inp_toks = count_tokens(prompt, model)
        out_toks = count_tokens(text, model)

    # --- 비용 계산 ---
    rates = PRICING.get(model)
    if rates is None:
        total_cost = None
    else:
        total_cost = inp_toks * rates["prompt"] + out_toks * rates["completion"]

    # --- 항상 4개 값 반환 ---
    return text, inp_toks, out_toks, total_cost


def get_simple_powerpoint_info(max_retries=3, delay_seconds=1):
    for attempt in range(1, max_retries + 1):
        try:
            # PowerPoint 애플리케이션에 연결
            ppt_app = win32com.client.GetObject(Class="PowerPoint.Application")
            
            if not ppt_app or not hasattr(ppt_app, 'ActivePresentation'):
                raise RuntimeError("PowerPoint가 실행 중이 아니거나 열린 프레젠테이션이 없습니다.")
            
            presentation = ppt_app.ActivePresentation
            
            file_name = presentation.Name
            slide_count = presentation.Slides.Count
            
            return {
                "파일 이름": file_name,
                "슬라이드 수": slide_count
            }
        
        except Exception as e:
            print(f"{attempt}번째 시도에서 오류 발생: {e}")

            if attempt < max_retries:
                time.sleep(delay_seconds)


def get_shape_type(shape_type):
    """Shape 유형 번호를 문자열로 변환"""
    shape_types = {
        1: "AutoShape", 
        2: "CallOut",
        3: "Chart",
        4: "Comment",
        5: "Freeform",
        6: "Group",
        7: "EmbeddedOLEObject",
        8: "FormControl",
        9: "Line",
        10: "LinkedOLEObject",
        11: "LinkedPicture",
        12: "OLEControl",
        13: "Picture",
        14: "Placeholder",
        15: "MediaObject", 
        16: "TextEffect",
        17: "TextBox",
        18: "Table",
        19: "SmartArt",
        20: "WebVideo",
        21: "ContentApp"
    }
    return shape_types.get(shape_type, f"Unknown Type ({shape_type})")

def get_placeholder_type(placeholder_type):
    """Placeholder 유형 번호를 문자열로 변환"""
    placeholder_types = {
        1: "Title",
        2: "Body",
        3: "CenterTitle",
        4: "SubTitle",
        5: "VerticalTitle",
        6: "VerticalBody",
        7: "Object",
        8: "Chart",
        9: "Table",
        10: "ClipArt",
        11: "OrgChart",
        12: "Media",
        13: "VerticalObject",
        14: "Picture",
        15: "Slide Number",
        16: "Header",
        17: "Footer",
        18: "Date",
        19: "VerticalTitle2",
        20: "VerticalBody2" 
    }
    return placeholder_types.get(placeholder_type, f"Unknown Placeholder ({placeholder_type})")

import win32com.client
import traceback  # 오류 추적을 위해 추가


def safe(obj, attr, default=None):
    """Safely get an attribute, returning default if an error occurs."""
    try:
        if obj is None:
            return default
        if hasattr(obj, attr):
            val = getattr(obj, attr, default)
            if val is None:
                return default
            return val
        return default
    except Exception:
        return default


def rgb_of(font):
    if font is None:
        return None
    rgb = None
    try:
        fill = safe(font, "Fill")
        fore = safe(fill, "ForeColor")
        temp = safe(fore, "RGB")
        if temp is not None:
            rgb = temp
    except Exception:
        pass
    if rgb is None:
        try:
            col = safe(font, "Color")
            temp = safe(col, "RGB")
            if temp is not None:
                rgb = temp
        except Exception:
            pass
    return rgb


def snap(font):
    if font is None:
        return (None, 0.0, False, False, False, None, False, False, False)
    size_val = safe(font, "Size", 0)
    try:
        size_f = float(size_val)
    except Exception:
        size_f = 0.0
    return (
        safe(font, "Name"),
        round(size_f, 1),
        bool(safe(font, "Bold", 0)),
        bool(safe(font, "Italic", 0)),
        bool(safe(font, "Underline", 0)),
        rgb_of(font),
        bool(safe(font, "Strikethrough", 0)),
        bool(safe(font, "Subscript", 0)),
        bool(safe(font, "Superscript", 0)),
    )


def make_run_dict(text_range_segment):
    if text_range_segment is None:
        return {"Text": "", "Font": {}}
    f = safe(text_range_segment, "Font")
    text = safe(text_range_segment, "Text", "")
    font_dict = {}
    if f:
        rgb = rgb_of(f)
        font_dict = {
            "Name": safe(f, "Name"),
            "Size": safe(f, "Size"),
            "Bold": bool(safe(f, "Bold", 0)),
            "Italic": bool(safe(f, "Italic", 0)),
            "Underline": bool(safe(f, "Underline", 0)),
            "Strikethrough": bool(safe(f, "Strikethrough", 0)),
            "Subscript": bool(safe(f, "Subscript", 0)),
            "Superscript": bool(safe(f, "Superscript", 0)),
        }
        if rgb is not None:
            try:
                font_dict["Color"] = {"R": rgb & 0xFF, "G": (rgb >> 8) & 0xFF, "B": (rgb >> 16) & 0xFF}
            except Exception:
                pass
    else:
        font_dict = {"Name": None, "Size": None, "Bold": False, "Italic": False,
                     "Underline": False, "Strikethrough": False, "Subscript": False,
                     "Superscript": False}

    hyperlink = None
    try:
        act = safe(safe(text_range_segment, "ActionSettings"), 1)
        hyperlink = safe(safe(act, "Hyperlink"), "Address")
    except Exception:
        pass

    return {"Text": text, "Font": font_dict, "Hyperlink": hyperlink} if hyperlink else {"Text": text, "Font": font_dict}    

def parse_paragraph_bullets(text_frame):
    """
    TextFrame에서 문단 단위의 bullet 정보를 추출
    - Bullet 여부
    - 들여쓰기 정보
    - Bullet Level (파워포인트 단락 레벨)
    """
    result = []

    if not safe(text_frame, "HasText", False):
        return result

    tr = safe(text_frame, "TextRange")
    if not tr:
        return result

    try:
        para_count = tr.Paragraphs().Count
    except Exception:
        return result

    for i in range(1, para_count + 1):
        try:
            p = tr.Paragraphs(i)
            pf = safe(p, "ParagraphFormat")
            bullet = safe(pf, "Bullet")

            has_bullet = bool(safe(bullet, "Visible", False))

            if has_bullet and bullet:
                para_info = {
                    "ParagraphIndex": i,
                    "Text": safe(p, "Text", "").rstrip("\r\n"),
                    "HasBullet": has_bullet,
                    "Level": safe(p, "Level"),
                }

                para_info.update({
                    "BulletCharacter": safe(bullet, "Character"),
                    "BulletType": safe(bullet, "Type"),
                    "BulletRelativeSize": safe(bullet, "RelativeSize"),
                    "BulletFontName": safe(safe(bullet, "Font"), "Name"),
                })

                # 들여쓰기(참고용)
                para_info["FirstLineIndent"] = safe(pf, "FirstLineIndent")
                para_info["LeftIndent"] = safe(pf, "LeftIndent")

                result.append(para_info)
        except Exception:
            continue

    return result

def parse_text_frame_debug(text_frame):
    out = {"Has Text": False}
    if not safe(text_frame, "HasText", False):
        return out
    tr = safe(text_frame, "TextRange")
    if not tr:
        return out
    full = safe(tr, "Text", "")
    out.update({"Has Text": True, "Text": full, "Runs": []})

    if not full:
        out["Paragraphs"] = []
        return out

    runs = []
    n = len(full)
    try:
        cur_idx = 1
        cur_snap = snap(safe(tr.Characters(cur_idx, 1), "Font"))
        for i in range(2, n + 1):
            nxt_snap = snap(safe(tr.Characters(i, 1), "Font"))
            if nxt_snap != cur_snap:
                seg_len = i - cur_idx
                if seg_len > 0:
                    runs.append(make_run_dict(tr.Characters(cur_idx, seg_len)))
                cur_idx = i
                cur_snap = nxt_snap
        last_len = n - cur_idx + 1
        if last_len > 0:
            runs.append(make_run_dict(tr.Characters(cur_idx, last_len)))
    except Exception as e:
        print(f"Error parsing runs: {e}")
        traceback.print_exc()
        runs.append(make_run_dict(tr))

    out["Runs"] = runs

    # 🔹 여기서 문단/글머리표 정보 추가
    out["Paragraphs"] = parse_paragraph_bullets(text_frame)

    return out




def parse_table(table):
    """테이블 정보 파싱 (결과를 dict로 반환)"""
    result = {}
    try:
        rows = getattr(table.Rows, "Count", 0)
        cols = getattr(table.Columns, "Count", 0)
        result["Dimensions"] = {"Rows": rows, "Columns": cols}
        result["FirstRow"]   = getattr(table, "FirstRow", None)
        result["LastRow"]    = getattr(table, "LastRow", None)
        result["FirstCol"]   = getattr(table, "FirstCol", None)
        result["LastCol"]    = getattr(table, "LastCol", None)

        # 샘플 셀 내용
        samples = {}
        max_r = min(3, rows)
        max_c = min(3, cols)
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                key = f"Cell({r},{c})"
                try:
                    txt = table.Cell(r, c).Shape.TextFrame.TextRange.Text
                    samples[key] = txt[:30] + ("..." if len(txt) > 30 else "")
                except Exception:
                    samples[key] = None
        result["Sample Cells"] = samples

    except Exception as e:
        result["Table Parsing Error"] = str(e)
    return result


def parse_chart(chart):
    """차트 정보 파싱 (결과를 dict로 반환)"""
    result = {}
    try:
        ct = getattr(chart, "ChartType", None)
        chart_types = {
            -4100: "xlColumnClustered", -4101: "xlColumnStacked",
            -4170: "xlBarClustered",    -4102: "xlLineStacked",
             73:    "xlPie"
        }
        result["Chart Type"] = chart_types.get(ct, f"Unknown ({ct})")
        result["Has Legend"] = bool(getattr(chart, "HasLegend", False))
        result["Has Title"]  = bool(getattr(chart, "HasTitle", False))
        if result["Has Title"]:
            result["Title Text"] = getattr(chart.ChartTitle, "Text", None)

        # 시리즈 정보
        series_info = {}
        try:
            sc = chart.SeriesCollection()
            count = getattr(sc, "Count", 0)
            series_info["Count"] = count
            for i in range(1, min(count, 4) + 1):
                try:
                    series_info[f"Series {i} Name"] = sc.Item(i).Name
                except Exception:
                    series_info[f"Series {i} Name"] = None
        except Exception as se:
            series_info["Error"] = str(se)
        result["Series"] = series_info

    except Exception as e:
        result["Chart Parsing Error"] = str(e)
    return result


# def parse_group_shapes(group_shape):
#     """그룹 내부 Shape 정보 파싱 (결과를 dict로 반환)"""
#     result = {}
#     try:
#         count = getattr(group_shape.GroupItems, "Count", 0)
#         result["Group Items Count"] = count
#         items = {}
#         for i in range(1, min(count, 3) + 1):
#             try:
#                 sub = group_shape.GroupItems.Item(i)
#                 items[f"Item {i}"] = {
#                     "Name": getattr(sub, "Name", None),
#                     "Type": getattr(sub, "Type", None)
#                 }
#             except Exception:
#                 items[f"Item {i}"] = None
#         result["Items"] = items

#     except Exception as e:
#         result["Group Parsing Error"] = str(e)
#     return result

def parse_group_shapes(group_shape):
    """
    Group 내부의 모든 Shape을 재귀적으로 파싱
    """
    result = []

    try:
        group_items = group_shape.GroupItems
        count = group_items.Count

        for i in range(1, count + 1):
            sub = group_items.Item(i)

            sid = sub.Id
            name = sub.Name
            stype = sub.Type
            left = sub.Left
            top = sub.Top
            width = sub.Width
            height = sub.Height

            item_info = {
                "Shape_Id": sid,
                "Name": name,
                "Type": SHAPE_TYPE_MAP.get(stype, stype),
                "Position_Left": left,
                "Position_Top": top,
                "Size_Width": width,
                "Size_Height": height,
            }

            # ---- 텍스트 ----
            if sub.HasTextFrame: # text
                tf = sub.TextFrame
                if tf.HasText:
                    item_info["Text"] = extract_text_from_shape(sub)

            if stype == 6:  # msoGroup
                item_info["GroupItems"] = parse_group_shapes(sub)

            elif stype == 13:  # Picture
                item_info["Picture"] = {
                    "AlternativeText": sub.AlternativeText
                }

            elif stype == 3:  # Chart
                chart = sub.Chart
                item_info["Chart"] = {
                    "ChartType": chart.ChartType,
                    "HasTitle": chart.HasTitle
                }

            elif stype == 19:  # Table
                table = sub.Table
                item_info["Table"] = {
                    "Rows": table.Rows.Count,
                    "Columns": table.Columns.Count
                }

            result.append(item_info)

    except Exception as e:
        return {"Group Parsing Error": str(e)}

    return result



def parse_picture(picture):
    """이미지 정보 파싱 (결과를 dict로 반환)"""
    result = {}
    try:
        result["Type"] = getattr(picture, "Type", None)
        result["Scale"] = {
            "Width %": getattr(picture, "ScaleWidth", None),
            "Height %": getattr(picture, "ScaleHeight", None)
        }
        pf = getattr(picture, "PictureFormat", None)
        if pf:
            pic_fmt = {}
            for attr in ("Brightness", "Contrast"):
                if hasattr(pf, attr):
                    pic_fmt[attr] = getattr(pf, attr)
            crop = getattr(pf, "Crop", None)
            if crop:
                pic_fmt["Crop"] = {
                    "Left": getattr(crop, "ShapeLeft", None),
                    "Top": getattr(crop, "ShapeTop", None),
                    "Width": getattr(crop, "ShapeWidth", None),
                    "Height": getattr(crop, "ShapeHeight", None)
                }
            result["PictureFormat"] = pic_fmt

    except Exception as e:
        result["Picture Parsing Error"] = str(e)
    return result


def parse_placeholder_details(placeholder):
    """Placeholder 상세 정보 파싱 (결과를 dict로 반환)"""
    result = {}
    try:
        pf = placeholder.PlaceholderFormat
        ptype = getattr(pf, "Type", None)
        result["Placeholder Type"]  = ptype
        result["Placeholder Type Name"] = get_placeholder_type(ptype)
        result["Placeholder ID"]    = getattr(placeholder, "Id", None)
        result["Placeholder Index"] = getattr(pf, "Index", None)
        if hasattr(pf, "ContainedType"):
            result["Contained Type"] = getattr(pf, "ContainedType", None)

    except Exception as e:
        result["Placeholder Parsing Error"] = str(e)
    return result

def parse_shape_details(shape):
    """Shape 유형별 세부 정보 파싱 (결과를 dict로 반환)"""
    result = {}

    # 공통 속성
    #try:
    result["Visibility"] = "Visible" if getattr(shape, "Visible", False) else "Hidden"
    result["Z-Order"]    = getattr(shape, "ZOrderPosition", None)
    if hasattr(shape, "Rotation"):
        result["Rotation (°)"] = getattr(shape, "Rotation", None)
    result["ID"] = getattr(shape, "Id", None)

    # 투명도
    fill = getattr(shape, "Fill", None)
    if fill and hasattr(fill, "Transparency"):
        result["Fill Transparency (%)"] = fill.Transparency * 100

    # 선 정보
    line = getattr(shape, "Line", None)
    if line and getattr(line, "Visible", False):
        line_info = {
            "Width (pt)": getattr(line, "Weight", None)
        }
        fore = getattr(line, "ForeColor", None)
        if fore and hasattr(fore, "RGB"):
            rgb = fore.RGB
            line_info["Color"] = {
                "R": rgb & 0xFF,
                "G": (rgb >> 8) & 0xFF,
                "B": (rgb >> 16) & 0xFF
            }
        result["Line"] = line_info

    #except Exception as e:
        #result["General Properties Error"] = str(e)

    # 타입별 세부 정보
    #try:
    t = getattr(shape, "Type", None)
    tf = None
    if safe(shape, "HasTextFrame", False):
        tf = shape.TextFrame
    elif safe(shape, "HasTextFrame2", False):
        tf = shape.TextFrame2

    if tf is None or not safe(tf, "HasText", False):
        print("No text")

    # debug 파싱 함수로부터 runs 정보 얻기
    parsed = parse_text_frame_debug(tf)
    runs = parsed.get("Runs", [])
    paras = parsed.get("Paragraphs", [])

    # result["TextFrame"] 에 run 별로 저장
    result["TextFrame"] = {
    "FullText": parsed.get("Text", ""),
    "Runs": [],
    "Paragraphs": paras,   
}
    for idx, run in enumerate(runs, start=1):
        # 원하는 속성만 골라 담거나, run 전체를 그대로 담을 수도 있습니다.
        run_info = {
            "RunIndex": idx,
            "Text": run.get("Text", ""),
            "Font": run.get("Font"),
            "Hyperlink": run.get("Hyperlink")
        }
        result["TextFrame"]["Runs"].append(run_info)
    

    # Placeholder
    if t == 14:
        print("placeholder!!!")
        result["Placeholder"] = parse_placeholder_details(shape)

    # Group
    elif t == 6:
        result["GroupShapes"] = parse_group_shapes(shape)

    # Table
    elif t == 18:
        result["Table"] = parse_table(shape.Table)

    # Chart
    elif t == 3:
        result["Chart"] = parse_chart(shape.Chart)

    # Picture
    elif t in (11, 13):
        result["Picture"] = parse_picture(shape)

    # SmartArt
    elif t == 19:
        result["SmartArt Nodes"] = getattr(shape.SmartArt.AllNodes, "Count", None)

    # OLE Object
    elif t in (7, 10):
        prog = getattr(shape.OLEFormat, "ProgID", None) if hasattr(shape, "OLEFormat") else None
        result["OLE Class Type"] = prog or "Unknown"

    # Media
    elif t == 15:
        result["Media Type"] = getattr(shape, "MediaType", "Unknown")

    #except Exception as e:
     #   result["Shape Detail Error"] = str(e)

    return result


def parse_slide_notes(slide):
    """슬라이드 노트 파싱 (결과를 dict로 반환)"""
    result = {}
    try:
        # 노트 페이지 유무
        has_notes = getattr(slide, "HasNotesPage", False)
        result["Has Notes Page"] = bool(has_notes)

        if has_notes:
            notes_page = slide.NotesPage
            shapes = notes_page.Shapes
            count = getattr(shapes, "Count", 0)
            result["Notes Shapes Count"] = count

            # 노트 텍스트 수집
            texts = []
            for i in range(1, count + 1):
                shape = shapes(i)
                ph = getattr(shape, "PlaceholderFormat", None)
                if ph and getattr(ph, "Type", None) == 2:
                    tf = getattr(shape, "TextFrame", None)
                    if tf and getattr(tf, "HasText", False):
                        texts.append(shape.TextFrame.TextRange.Text)

            # 내용 유무에 따라 설정
            if texts:
                result["Notes Content"] = "".join(texts)
            else:
                result["Notes Content"] = None
        else:
            result["Notes Content"] = None

    except Exception as e:
        result["Error parsing notes"] = str(e)

    return result


def parse_slide_properties(slide):
    """슬라이드 속성 파싱 (결과를 dict로 반환)"""
    result = {}
    try:
        # Layout 은 COM 상에서 단순 enum(int) 이므로
        # .Type/.Name 을 호출하면 int 에서 에러가 남.
        # 대신 코드값만 저장하고, CustomLayout 객체를 쓰세요.

        layout_code = getattr(slide, "Layout", None)
        if layout_code is not None:
            result["Slide Layout Code"] = layout_code

        # CustomLayout 은 객체이므로 이름/인덱스 등을 가져올 수 있음
        custom = getattr(slide, "CustomLayout", None)
        if custom is not None:
            result["CustomLayout Name"]  = getattr(custom, "Name", None)
            result["CustomLayout Index"] = getattr(custom, "Index", None)

        # 배경 채우기 정보
        bg = getattr(slide, "Background", None)
        if bg is not None:
            fill = getattr(bg, "Fill", None)
            if fill is not None:
                # fill.Type 은 안전하게 getattr 으로
                t = getattr(fill, "Type", None)
                fill_types = {1: "Solid", 2: "Pattern", 3: "Gradient", 4: "Texture", 5: "Picture"}
                result["Background Fill Type"] = fill_types.get(t, f"Unknown ({t})")

        # 전환 효과
        trans = getattr(slide, "SlideShowTransition", None)
        if trans is not None:
            result["Transition Effect"]   = getattr(trans, "EntryEffect", "None")
            result["Advance Time (s)"]    = getattr(trans, "AdvanceTime", "Manual")
            result["Advance On Click"]    = bool(getattr(trans, "AdvanceOnClick", False))
            result["Advance On Time"]     = bool(getattr(trans, "AdvanceOnTime", False))

    except Exception as e:
        result["error"] = str(e)

    return result



# def parse_active_slide_objects(slide_num:int=1):
#     """슬라이드 객체 파싱 메인 함수"""
#     output = {} # 출력을 저장할 문자열 초기화
    
#     try:
#         # Connect to running PowerPoint instance
#         ppt = win32com.client.GetObject(Class="PowerPoint.Application")
        
#         # Get active presentation
#         presentation = ppt.ActivePresentation
        
#         # Check if there is an active presentation
#         if not presentation:
#             output['status'] = "No active presentation found."
#             return output['status']
        
#         # 프레젠테이션 정보 추가
#         output["Presentation_Name"] = f"{presentation.Name}"
#         output["Total_Slide_Number"] = f"{presentation.Slides.Count}"
        
#         # 슬라이드 범위 확인
#         if slide_num > presentation.Slides.Count or slide_num < 1:
#             output["status"] = f"Invalid slide number. Please provide a number between 1 and {presentation.Slides.Count}."
#             return output["status"]
        
#         # Access the specified slide
#         slide = presentation.Slides(slide_num)
        
#         # 슬라이드 속성 파싱
#         output["Slide_Properties"] = parse_slide_properties(slide)
        
#         # Get the number of shapes in the slide
#         shape_count = slide.Shapes.Count
#         output["Objects_Overview"] = f"Found {shape_count} objects in slide number {slide_num}."
#         output["Objects_Detail"] = []

#         # Iterate through each shape
#         for i in range(1, shape_count + 1):
#             shape = slide.Shapes(i)
#             shape_info = {
#                 "Object_number": i,
#                 "Shape_Id": shape.Id,
#                 "Name": shape.Name,
#                 "Type": get_shape_type(shape.Type),
#                 "Position_Left": shape.Left,
#                 "Position_Top": shape.Top,
#                 "Size_Width": shape.Width,
#                 "Size_Height": shape.Height,
#                 "More_detail": parse_shape_details(shape),
                
#             }
#             output["Objects_Detail"].append(shape_info)
        
#         # 슬라이드 노트 파싱
#         output["Slide_Notes"] = parse_slide_notes(slide)
#     except pywintypes.com_error as e:
#         output["Error"] = f"COM error: {e}"
#     # except Exception as e:
#     #     output["Error"] = f"Error: {e}"

#     return output

def parse_active_slide_objects(slide_num: int, prs_obj):
    """
    Parse Every Object Information from a Slide.
    Args:
        slide_num (int): 파싱할 슬라이드 번호 (1-based)
        prs_obj: PPTContainer.prs 또는 win32com Presentation 객체
    """

    output = {}

    try:
        presentation = prs_obj

        if not presentation:
            return {"status": "No active presentation found."}

        slides = presentation.Slides
        slide_count = slides.Count

        output["Presentation_Name"] = presentation.Name
        output["Total_Slide_Number"] = slide_count
        output["Current_Slide_Number"] = slide_num

        output["Slide Width"]  = presentation.PageSetup.SlideWidth
        output["Slide Height"] = presentation.PageSetup.SlideHeight
        
        if slide_num < 1 or slide_num > slide_count:
            return {"status": f"Invalid slide number (1~{slide_count})"}

        slide = slides(slide_num)

        output["Slide_Properties"] = parse_slide_properties(slide)

        shapes = slide.Shapes
        shape_count = shapes.Count

        output["Objects_Overview"] = f"Found {shape_count} objects"
        output["Objects_Detail"] = []

        for i in range(1, shape_count + 1):
            shape = shapes(i)

            # ---- COM 속성 캐싱 ----
            sid = shape.Id
            name = shape.Name
            stype = shape.Type
            left = shape.Left
            top = shape.Top
            width = shape.Width
            height = shape.Height

            shape_info = {
                "Object_number": i,
                "Shape_Id": sid,
                "Name": name,
                "Type": SHAPE_TYPE_MAP.get(stype, stype),
                "Position_Left": left,
                "Position_Top": top,
                "Size_Width": width,
                "Size_Height": height,
                "More_detail": parse_shape_details_fast(shape, stype),
            }

            output["Objects_Detail"].append(shape_info)

        output["Slide_Notes"] = parse_slide_notes(slide)

    except pywintypes.com_error as e:
        output["Error"] = f"COM error: {e}"

    return output


def extract_text_from_shape(shape, indent_level=1):
    """
    모든 유형의 도형에서 텍스트를 추출 (결과를 dict로 반환)
    """
    result = {}
    # try:
        # 1) TextFrame 지원 객체
    if getattr(shape, "HasTextFrame", False) and shape.TextFrame.HasText:
        tr = shape.TextFrame.TextRange

        tf = None
        if safe(shape, "HasTextFrame", False):
            tf = shape.TextFrame
        elif safe(shape, "HasTextFrame2", False):
            tf = shape.TextFrame2

        if tf is None or not safe(tf, "HasText", False):
            print("No text")

        # 여기서 이미 FullText / Runs / Paragraphs(=bullet+Level 포함)를 모두 얻음
        parsed = parse_text_frame_debug(tf)

        # 그대로 넣되, Hyperlink만 살짝 추가
        result["TextFrame"] = parsed

        try:
            hl = tr.ActionSettings(1).Hyperlink
            addr = getattr(hl, "Address", None)
            if addr:
                result["TextFrame"]["Hyperlink"] = addr
        except:
            pass

    # 2) TextFrame2 (Office2007+) 지원 객체
    elif getattr(shape, "HasTextFrame2", False) and shape.TextFrame2.TextRange.Text:
        tr2 = shape.TextFrame2.TextRange
        font2 = tr2.Font
        result["TextFrame2"] = {
            "Text": getattr(tr2, "Text", ""),
            "Font": {
                "Name": getattr(font2, "Name", None),
                "Size": getattr(font2, "Size", None),
                "Bold": getattr(font2, "Bold", None),
                "Italic": getattr(font2, "Italic", None),
            }
        }

    # 3) Table 내 텍스트
    elif getattr(shape, "Type", None) == 19 and hasattr(shape, "Table"):
        tbl = shape.Table
        rows, cols = getattr(tbl.Rows, "Count", 0), getattr(tbl.Columns, "Count", 0)
        cells = {}
        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                key = f"Cell({r},{c})"
                try:
                    txt = tbl.Cell(r, c).Shape.TextFrame.TextRange.Text
                except:
                    txt = None
                cells[key] = txt
        result["TableText"] = {"Rows": rows, "Columns": cols, "Cells": cells}

    # 4) Chart 내 텍스트
    elif getattr(shape, "Type", None) == 3 and hasattr(shape, "Chart"):
        chart = shape.Chart
        chart_info = {
            "Title": getattr(chart.ChartTitle, "Text", None) if getattr(chart, "HasTitle", False) else None,
            "Axes": {}
        }
        if hasattr(chart, "Axes"):
            for grp in (1, 2, 3):
                for typ in (1, 2):
                    try:
                        ax = chart.Axes(grp, typ)
                        if getattr(ax, "HasTitle", False):
                            chart_info["Axes"][f"{grp},{typ}"] = ax.AxisTitle.Text
                    except:
                        pass
        result["ChartText"] = chart_info

    # 5) SmartArt 내 텍스트
    elif getattr(shape, "Type", None) == 24 and hasattr(shape, "SmartArt"):
        nodes = getattr(shape.SmartArt, "AllNodes", None)
        smart = {}
        if nodes:
            for i in range(1, getattr(nodes, "Count", 0) + 1):
                try:
                    txt = nodes.Item(i).TextFrame2.TextRange.Text
                except:
                    txt = None
                smart[f"Node {i}"] = txt
        result["SmartArtText"] = smart

    else:
        result["Text"] = None

    # except Exception as e:
    #     result["Error"] = str(e)

    return result



def get_alignment_type(alignment_val):
    # 단락 정렬 값을 텍스트로 변환
    alignment_types = {
        1: "Left",
        2: "Center",
        3: "Right",
        4: "Justify",
        5: "Distributed"
    }
    return alignment_types.get(alignment_val, f"Unknown Alignment ({alignment_val})")

def parse_group_shape(group_shape, indent_level=1):
    """Recursively parse all items within a group object"""
    output = ""  # 출력을 저장할 문자열 초기화
    try:
        indent = "  " * indent_level
        group_items_count = group_shape.GroupItems.Count
        output += f"{indent}Number of objects in group: {group_items_count}"
        
        # Iterate through each item in the group
        for j in range(1, group_items_count + 1):
            group_item = group_shape.GroupItems.Item(j)
            output += f"\n{indent}Object in group {j}:"
            output += f"\n{indent}  Name: {group_item.Name}"
            output += f"\n{indent}  Type: {get_shape_type(group_item.Type)}"
            output += f"\n{indent}  Position: Left={group_item.Left}, Top={group_item.Top}"
            output += f"\n{indent}  Size: Width={group_item.Width}, Height={group_item.Height}"
            
            # 그룹 내 개체의 텍스트 추출
            output += extract_text_from_shape(group_item, indent_level + 1)
            
            # Process recursively if the item in the group is another group
            if group_item.Type == 6:  # Group
                output += parse_group_shape(group_item, indent_level + 1)
            else:
                # Parse regular shape details
                output += parse_shape_details(group_item, indent_level + 1)
                
    except Exception as e:
        output += f"\n{indent}Group object parsing error: {e}"
    
    return output

def parse_shape_details(shape, indent_level=1):
    """
    Shape 유형별 세부 정보 파싱 (결과를 dict로 반환)
    """
    result = {}
    
    # 1) 먼저 텍스트 관련 정보를 dict로 가져와 병합
    text_info = extract_text_from_shape(shape, indent_level)
    if isinstance(text_info, dict):
        result.update(text_info)
    
    # 2) Shape 유형별 추가 정보
    try:
        t = getattr(shape, "Type", None)
        
        # 그룹
        if t == 6:
            grp = {"Group": {
                "Name": getattr(shape, "Name", None),
                "Items": parse_group_shapes(shape)  # 이 함수도 dict 반환 가정
            }}
            result.update(grp)
        
        # 그림(Picture)
        elif t == 13:
            pic_info = {"Picture": {
                "Name": getattr(shape, "Name", None),
                "AlternativeText": getattr(shape, "AlternativeText", None)
            }}
            result.update(pic_info)
        
        # 차트(Chart)
        elif t == 3:
            chart = getattr(shape, "Chart", None)
            chart_info = {"Chart": {
                "Name": getattr(shape, "Name", None),
                "ChartType": getattr(chart, "ChartType", None) if chart else None,
                "HasTitle": getattr(chart, "HasTitle", None) if chart else None
            }}
            result.update(chart_info)
        
        # 테이블(Table)
        elif t == 19:
            table = getattr(shape, "Table", None)
            table_info = {"Table": {
                "Name": getattr(shape, "Name", None),
                "Rows": getattr(table.Rows, "Count", None) if table else None,
                "Columns": getattr(table.Columns, "Count", None) if table else None
            }}
            result.update(table_info)
    
    except Exception as e:
        result["Shape Detail Error"] = str(e)
    
    return result

def parse_shape_details_fast(shape, stype):
    result = {}

    # ---- 텍스트가 있을 때만 파싱 (4번) ----
    if shape.HasTextFrame:
        tf = shape.TextFrame
        if tf.HasText:
            result["Text"] = extract_text_from_shape(shape)

    try:
        # Group
        if stype == 6:
            result["Group"] = {
                "Items": parse_group_shapes(shape)
            }

        # Picture
        elif stype == 13:
            result["Picture"] = {
                "AlternativeText": shape.AlternativeText
            }

        # Chart
        elif stype == 3:
            chart = shape.Chart
            result["Chart"] = {
                "ChartType": chart.ChartType,
                "HasTitle": chart.HasTitle
            }

        # Table
        elif stype == 19:
            table = shape.Table
            result["Table"] = {
                "Rows": table.Rows.Count,
                "Columns": table.Columns.Count
            }

    except Exception as e:
        result["Shape Detail Error"] = str(e)

    return result


# output = parse_active_slide_objects()
# print(output)