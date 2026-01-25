import win32com.client
import json
import re
import time
import os

from llm_client import call_llm
from utils import parse_llm_response
from prompts import STYLE_MAPPING_PROMPT

import logging
logger = logging.getLogger(__name__)

# --- Internal Helper Functions ---

def _hex_to_rgb_int(hex_str):
    """Converts HEX string (#FFFFFF or FFFFFF) to win32-compatible BGR integer."""
    hex_str = hex_str.lstrip('#')
    if len(hex_str) != 6:
        raise ValueError("HEX code must be 6 characters long (e.g., FF0000).")
    
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    return (b << 16) | (g << 8) | r


def _find_shape_by_id(prs, slide_number, shape_id):
    """Finds a specific Shape object by its unique ID on a given slide."""
    try:
        slide = prs.Slides(slide_number)
        for shape in slide.Shapes:
            if shape.Id == shape_id:
                return shape
    except Exception as e:
        raise ValueError(f"Error accessing slide {slide_number}: {e}")
    raise ValueError(f"Shape with ID {shape_id} not found on slide {slide_number}.")




def _get_text_runs_from_shape(shape):
    if not shape.HasTextFrame or not shape.TextFrame.HasText:
        return []
    return shape.TextFrame.TextRange.Runs()


def _get_text_runs_from_table_cell(shape, row_index, col_index):
    cell = shape.Table.Cell(row_index, col_index)
    if not cell.Shape.TextFrame.HasText:
        return []
    return cell.Shape.TextFrame.TextRange.Runs()


def _resolve_runs(
    prs,
    slide_number,
    shape_id,
    container="shape",
    row_index=None,
    col_index=None
):
    shape = _find_shape_by_id(prs, slide_number, shape_id)

    if container == "shape":
        return _get_text_runs_from_shape(shape)

    if container == "table_cell":
        if row_index is None or col_index is None:
            raise ValueError("row_index and col_index required for table_cell.")
        return _get_text_runs_from_table_cell(shape, row_index, col_index)

    raise ValueError(f"Unknown container type: {container}")



def _split_run_by_range(text_range, start, end):
    """
    Split a TextRange into [before][target][after] and
    return the target TextRange.
    """
    full_len = len(text_range.Text)

    if start < 0 or end > full_len or start >= end:
        raise ValueError("Invalid char range inside run.")

    # before
    if start > 0:
        text_range.Characters(1, start).InsertAfter(
            text_range.Characters(start + 1, full_len - start).Text
        )
        target = text_range.Characters(1, end - start)
    else:
        target = text_range.Characters(1, end)

    # after 제거
    if end < full_len:
        target.Characters(end - start + 1, full_len - end).Delete()

    return target



def undo_action(reason: str, container):
    """
    Rollback PPT to previous backup
    """
    ppt_app = container.prs.Application
    container.prs.Close()
    time.sleep(0.5)
    container.prs = ppt_app.Presentations.Open(os.path.abspath(container.backup_path))


#########################################################################
######################## [A]  Text Style Editing ########################
#########################################################################
def _get_text_with_offsets(
    prs,
    slide_number: int,
    shape_id: int,
    *,
    container: str = "shape",
    row_index: int = None,
    col_index: int = None,
):
    shape = _find_shape_by_id(prs, slide_number, shape_id)

    if container == "shape":
        if not shape.HasTextFrame or not shape.TextFrame.HasText:
            return "", []
        tr = shape.TextFrame.TextRange

    elif container == "table_cell":
        if row_index is None or col_index is None:
            raise ValueError("row_index / col_index required.")
        cell = shape.Table.Cell(row_index, col_index)
        if not cell.Shape.TextFrame.HasText:
            return "", []
        tr = cell.Shape.TextFrame.TextRange

    else:
        raise ValueError(f"Unknown container: {container}")

    text = tr.Text or ""
    return text, list(range(len(text)))


def _normalize_char_range(
    text: str,
    char_start_index: int,
    target_text: str,
    char_end: int = None,
):
    if char_start_index < 0 or char_start_index >= len(text):
        raise ValueError("char_start_index out of range.")

    expected_len = len(target_text)
    primary_end = char_start_index + expected_len

    if text[char_start_index:primary_end] == target_text:
        return char_start_index, primary_end

    # If llm gave end
    if char_end is not None:
        if text[char_start_index:char_end] == target_text:
            return char_start_index, char_end

    # Window search
    window_start = max(0, char_start_index - 5)
    window_end = min(len(text), char_start_index + expected_len + 5)
    window = text[window_start:window_end]

    idx = window.find(target_text)
    if idx != -1:
        start = window_start + idx
        return start, start + expected_len

    raise ValueError("Unable to resolve exact character range.")

def _get_runs_by_shape_id(slide_json: dict, shape_id: int):
    for obj in slide_json.get("Objects_Detail", []):
        if obj.get("Shape_Id") == shape_id:
            return (
                obj.get("More_detail", {})
                   .get("Text", {})
                   .get("TextFrame", {})
                   .get("Runs", [])
            )
    raise ValueError(f"Shape_Id {shape_id} not found in slide JSON.")

def _iter_run_slices_from_shape_json(
    slide_json: dict,
    shape_id: int,
    start: int,
    end: int,
):
    runs = _get_runs_by_shape_id(slide_json, shape_id)

    for run in runs:
        rs = run["Run_Start_Index"]
        re = rs + len(run["Text"])

        s = max(start, rs)
        e = min(end, re)

        if s < e:
            yield s, e, run

def _apply_font_snapshot(font, snap: dict):
    if "Name" in snap: font.Name = snap["Name"]
    if "Size" in snap: font.Size = snap["Size"]

    font.Bold = int(snap.get("Bold", False))
    font.Italic = int(snap.get("Italic", False))
    font.Underline = int(snap.get("Underline", False))

    if hasattr(font, "Strikethrough"): font.Strikethrough = int(snap.get("Strikethrough", False))
    if hasattr(font, "Subscript"): font.Subscript = int(snap.get("Subscript", False))
    if hasattr(font, "Superscript"): font.Superscript = int(snap.get("Superscript", False))

    color = snap.get("Color")
    if color:
        font.Color.RGB = (color["B"] << 16) | (color["G"] << 8) | color["R"]

def _apply_overrides(font, *,
                    font_name=None,
                    font_size=None,
                    bold=None,
                    italic=None,
                    underline=None,
                    color_hex=None):
    if font_name is not None:
        font.Name = font_name
    if font_size is not None:
        font.Size = font_size
    if bold is not None:
        font.Bold = int(bold)
    if italic is not None:
        font.Italic = int(italic)
    if underline is not None:
        font.Underline = int(underline)
    if color_hex is not None:
        font.Color.RGB = _hex_to_rgb_int(color_hex)

def set_text_style_preserve_runs(
    prs,
    slide_number: int,
    shape_id: int,
    char_start_index: int,
    target_text: str,
    slide_json: dict,
    *,
    char_end: int = None,
    container: str = "shape",
    row_index: int = None,
    col_index: int = None,
    font_name: str = None,
    font_size=None,
    bold=None,
    italic=None,
    underline=None,
    color_hex=None,
):
    # 1. 정확한 char range 보정
    text, _ = _get_text_with_offsets(
        prs,
        slide_number,
        shape_id,
        container=container,
        row_index=row_index,
        col_index=col_index,
    )

    start, end = _normalize_char_range(
        text=text,
        char_start_index=char_start_index,
        target_text=target_text,
        char_end=char_end,
    )

    # 2. TextRange 확보
    shape = _find_shape_by_id(prs, slide_number, shape_id)
    tr = (
        shape.TextFrame.TextRange
        if container == "shape"
        else shape.Table.Cell(row_index, col_index).Shape.TextFrame.TextRange
    )

    # 3. run-level slice 순회
    for s, e, run in _iter_run_slices_from_shape_json(
        slide_json, shape_id, start, end
    ):
        length = e - s
        target = tr.Characters(s + 1, length)
        font = target.Font

        # 3-1. 기존 run font 복원
        _apply_font_snapshot(font, run["Font"])

        # 3-2. 요청된 속성만 override
        _apply_overrides(
            font,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            underline=underline,
            color_hex=color_hex,
        )

    return {
        "operation": "set_text_style_preserve_runs",
        "applied_range": [start, end],
        "shape_id": shape_id,
        "slide": slide_number,
    }




def _resolve_insert_position(
    text: str,
    preceding_text: str,
    char_start_index: int,
):
    if preceding_text:
        idx = text.find(preceding_text)
        if idx != -1:
            return idx + len(preceding_text)

    # fallback: clamp index
    return max(0, min(len(text), char_start_index))

# def insert_text_from_textbox(
#     prs,
#     slide_number,
#     shape_id,
#     preceding_text,
#     char_start_index,
#     new_text,
#     *,
#     container="shape",
#     row_index=None,
#     col_index=None,
# ):
#     text, _ = _get_text_with_offsets(
#         prs, slide_number, shape_id,
#         container=container,
#         row_index=row_index,
#         col_index=col_index,
#     )

#     insert_pos = _resolve_insert_position(
#         text=text,
#         preceding_text=preceding_text,
#         char_start_index=char_start_index,
#     )

#     shape = _find_shape_by_id(prs, slide_number, shape_id)
#     tr = (
#         shape.TextFrame.TextRange
#         if container == "shape"
#         else shape.Table.Cell(row_index, col_index).Shape.TextFrame.TextRange
#     )

#     # PowerPoint is 1-based
#     anchor = tr.Characters(insert_pos + 1, 0)

#     # style is inherited from anchor
#     anchor.InsertAfter(new_text)

#     return {
#         "operation": "insert",
#         "insert_pos": insert_pos,
#         "new_text": new_text,
#         "shape_id": shape_id,
#         "slide": slide_number,
#     }




# def delete_text_from_textbox(
#     prs,
#     slide_number,
#     shape_id,
#     target_text,
#     char_start_index,
#     *,
#     container="shape",
#     row_index=None,
#     col_index=None,
# ):
#     text, _ = _get_text_with_offsets(
#         prs, slide_number, shape_id,
#         container=container,
#         row_index=row_index,
#         col_index=col_index,
#     )

#     start, end = _normalize_char_range(
#         text=text,
#         char_start_index=char_start_index,
#         target_text=target_text,
#     )

#     shape = _find_shape_by_id(prs, slide_number, shape_id)
#     tr = (
#         shape.TextFrame.TextRange
#         if container == "shape"
#         else shape.Table.Cell(row_index, col_index).Shape.TextFrame.TextRange
#     )

#     tr.Characters(start + 1, end - start).Delete()

#     return {
#         "operation": "delete",
#         "range": [start, end],
#         "deleted_text": target_text,
#         "shape_id": shape_id,
#         "slide": slide_number,
#     }



def _capture_font_style(char_range):
    font = char_range.Font
    return {
        "Name": font.Name,
        "Size": font.Size,
        "Bold": font.Bold,
        "Italic": font.Italic,
        "Underline": font.Underline,
        "Color": font.Color.RGB,
    }

def _apply_font_style(char_range, style):
    font = char_range.Font
    font.Name = style["Name"]
    font.Size = style["Size"]
    font.Bold = style["Bold"]
    font.Italic = style["Italic"]
    font.Underline = style["Underline"]
    font.Color.RGB = style["Color"]


# def replace_text_from_textbox(
#     prs,
#     slide_number,
#     shape_id,
#     target_text,
#     char_start_index,
#     new_text,
#     slide_json,
#     *,
#     container="shape",
#     row_index=None,
#     col_index=None,
# ):
#     # 1. 기존 텍스트와 정규화
#     text, _ = _get_text_with_offsets(
#         prs, slide_number, shape_id,
#         container=container,
#         row_index=row_index,
#         col_index=col_index,
#     )

#     start, end = _normalize_char_range(
#         text=text,
#         char_start_index=char_start_index,
#         target_text=target_text,
#     )

#     runs = _get_runs_by_shape_id(slide_json, shape_id)

#     # 2. Shape / TextRange 접근
#     slide = prs.Slides(slide_number)
#     shape = next((s for s in slide.Shapes if s.Id == shape_id), None)
#     if not shape:
#         raise ValueError(f"Shape {shape_id} not found in slide {slide_number}")

#     tr = (
#         shape.TextFrame.TextRange
#         if container == "shape"
#         else shape.Table.Cell(row_index, col_index).Shape.TextFrame.TextRange
#     )

#     old_range = tr.Characters(start + 1, end - start)

#     # 3. 단일 Run인지 확인
#     runs_in_range = [
#         r for r in runs
#         if r['Run_Start_Index'] <= start < r['Run_Start_Index'] + len(r['Text'])
#     ]
#     single_run = len(runs_in_range) == 1 and (start + len(target_text) <= runs_in_range[0]['Run_Start_Index'] + len(runs_in_range[0]['Text']))

#     if single_run:
#         # 단일 Run: 스타일 보존 후 교체
#         style = _capture_font_style(old_range)
#         old_range.Text = new_text
#         new_range = tr.Characters(start + 1, len(new_text))
#         _apply_font_style(new_range, style)

#         # # AutoShape 적용
#         # if hasattr(shape.TextFrame, "AutoSize"):
#         #     shape.TextFrame.AutoSize = 1  # ppAutoSizeShapeToFitText

#         return {
#             "operation": "replace_single_run",
#             "slide": slide_number,
#             "shape_id": shape_id,
#             "new_text": new_text,
#         }

#     # 4. 다중 Run: 기존 LLM 처리
#     llm_prompt = [
#         {"role": "system", "content": STYLE_MAPPING_PROMPT},
#         {
#             "role": "user",
#             "content": json.dumps({
#                 "old_runs": runs,
#                 "target_text": text[start:end],
#                 "new_text": new_text
#             }, ensure_ascii=False)
#         }
#     ]

#     raw_response = call_llm(model='gpt-4.1', messages=llm_prompt)
#     response = raw_response.output[0].content[0].text
#     new_runs = parse_llm_response(response)
#     new_runs = new_runs[0] if isinstance(new_runs, list) and len(new_runs) > 0 else new_runs

#     old_range.Text = ""
#     current_range = old_range

#     for run_info in new_runs:
#         if not isinstance(run_info, dict): 
#             continue
#         text_seg = run_info.get("Text", "")
#         if not text_seg: 
#             continue
#         new_run_range = current_range.InsertAfter(text_seg)
#         font_info = run_info.get("Font", {})
#         f = new_run_range.Font
#         if font_info.get("Name"): f.Name = font_info["Name"]
#         if font_info.get("Size"): f.Size = font_info["Size"]
#         f.Bold = -1 if font_info.get("Bold") else 0
#         f.Italic = -1 if font_info.get("Italic") else 0
#         f.Underline = -1 if font_info.get("Underline") else 0
#         color = font_info.get("Color")
#         if color and all(k in color for k in ("R","G","B")):
#             f.Color.RGB = color["R"] + (color["G"] << 8) + (color["B"] << 16)
#         current_range = new_run_range

#     # AutoShape 적용
#     if hasattr(shape.TextFrame, "AutoSize"):
#         shape.TextFrame.AutoSize = 1

#     return {
#         "operation": "replace_runs_llm",
#         "slide": slide_number,
#         "shape_id": shape_id,
#         "new_runs": new_runs,
#     }


def replace_shape_text(
    prs,
    slide_number,
    shape_id,
    new_text,
    slide_json,
    agent_request,
    *,
    container="shape",
    row_index=None,
    col_index=None,
):
    """
    Replace the text of a PowerPoint shape while preserving run-level styles.
    If text overflows after replacement, shrink font sizes proportionally
    based on the original TextRange font size.
    """

    # 1. Load old run info
    old_runs = _get_runs_by_shape_id(slide_json, shape_id)

    # ---- (A) extract base font size from old runs (upper bound) ----
    old_base_font_size = None
    sizes = []
    for run in old_runs:
        size = run.get("Font", {}).get("Size")
        if size:
            sizes.append(size)

    if sizes:
        old_base_font_size = max(sizes)  # cap

    # 2. Resolve slide / shape / TextRange
    slide = prs.Slides(slide_number)
    shape = next((s for s in slide.Shapes if s.Id == shape_id), None)
    if not shape:
        raise ValueError(f"Shape {shape_id} not found in slide {slide_number}")

    tr = (
        shape.TextFrame.TextRange
        if container == "shape"
        else shape.Table.Cell(row_index, col_index).Shape.TextFrame.TextRange
    )

    # 3. LLM call (unchanged)
    task_description, action_type, slide_contents = agent_request
    llm_prompt = [
        {"role": "system", "content": STYLE_MAPPING_PROMPT},
        {
            "role": "user",
            "content": json.dumps(
                {
                    "user_request": {
                        "task_description": task_description,
                        "action_type": action_type,
                        "slide_contents": slide_contents,
                    },
                    "old_runs": old_runs,
                    "new_text": new_text,
                },
                ensure_ascii=False,
            ),
        },
    ]

    raw_response = call_llm(model="gpt-4.1", messages=llm_prompt)
    response_text = raw_response.output[0].content[0].text
    new_runs = parse_llm_response(response_text)
    new_runs = new_runs[0]  # unwrap

    # 4. Clear existing text (layout reset)
    tr.Text = ""
    current_range = tr

    # 5. Apply new runs
    for run in new_runs:
        if not isinstance(run, dict):
            continue

        text_seg = run.get("Text", "")
        if not text_seg:
            continue

        new_range = current_range.InsertAfter(text_seg)
        font_info = run.get("Font", {})
        f = new_range.Font

        if font_info.get("Name"):
            f.Name = font_info["Name"]
        if font_info.get("Size"):
            f.Size = font_info["Size"]

        f.Bold = -1 if font_info.get("Bold") else 0
        f.Italic = -1 if font_info.get("Italic") else 0
        f.Underline = -1 if font_info.get("Underline") else 0
        f.Subscript = -1 if font_info.get("Subscript") else 0
        f.Superscript = -1 if font_info.get("Superscript") else 0

        color = font_info.get("Color")
        if color and all(k in color for k in ("R", "G", "B")):
            f.Color.RGB = color["R"] + (color["G"] << 8) + (color["B"] << 16)

        current_range = new_range

    # 6. Manual shrink-on-overflow
    tf = shape.TextFrame
    tf.WordWrap = True
    tf.AutoSize = 0

    new_tr = tf.TextRange

    # ---- (B) overflow only → proportional shrink ----
    if old_base_font_size and new_tr.Length > 0:
        # 현재 최대 font size
        current_max_size = new_tr.Font.Size

        # overflow 발생 시에만 shrink
        if new_tr.BoundHeight > shape.Height:
            # shrink target size 계산
            target_size = min(current_max_size, old_base_font_size)

            # 1차: 전체 범위 기준 shrink
            new_tr.Font.Size = target_size

            # 여전히 overflow면 step shrink
            while new_tr.BoundHeight > shape.Height and new_tr.Font.Size > 6:
                new_tr.Font.Size -= 0.5

            # ---- (C) run-level 비율 유지 ----
            scale = new_tr.Font.Size / target_size if target_size else 1.0

            for i in range(1, new_tr.Length + 1):
                ch = new_tr.Characters(i, 1)
                if ch.Font.Size:
                    ch.Font.Size *= scale

    return {
        "operation": "replace_shape_text",
        "slide": slide_number,
        "shape_id": shape_id,
        "new_text": new_text,
        "new_runs": new_runs,
    }

# def replace_text_from_textbox(
#     prs,
#     slide_number,
#     shape_id,
#     target_text,
#     char_start_index,
#     new_text,
#     slide_json,
#     *,
#     container="shape",
#     row_index=None,
#     col_index=None,
# ):
#     #  기존 텍스트 + char_start_index 정규화
#     text, _ = _get_text_with_offsets(
#         prs, slide_number, shape_id,
#         container=container,
#         row_index=row_index,
#         col_index=col_index,
#     )

#     start, end = _normalize_char_range(
#         text=text,
#         char_start_index=char_start_index,
#         target_text=target_text,
#     )

#     runs = _get_runs_by_shape_id(slide_json, shape_id)

#     # Shape / TextRange 접근
#     slide = prs.Slides(slide_number)
#     shape = None
#     for s in slide.Shapes:
#         if s.Id == shape_id:
#             shape = s
#             break
#     if not shape:
#         raise ValueError(f"Shape {shape_id} not found in slide {slide_number}")

#     if container == "shape":
#         tr = shape.TextFrame.TextRange
#     elif container == "table":
#         tr = shape.Table.Cell(row_index, col_index).Shape.TextFrame.TextRange
#     else:
#         raise ValueError(f"Unknown container {container}")

#     # 적용 범위 Run-level 수집
#     old_range = tr.Characters(start + 1, end - start)

#     llm_prompt = [
#         {"role": "system", "content": STYLE_MAPPING_PROMPT},
#         {
#             "role": "user",
#             "content": json.dumps({
#                 "old_runs": runs,
#                 "target_text": text[start:end],
#                 "new_text": new_text
#             }, ensure_ascii=False)
#         }
#     ]

#     raw_response = call_llm(model='gpt-4.1', messages=llm_prompt)
#     # response_json = response.model_dump_json(indent=2)

#     # with open("replace_text_llm_raw.json", "w", encoding="utf-8") as f:
#     #     f.write(response_json)

    
#     # 1. LLM 응답에서 텍스트 추출
#     response = raw_response.output[0].content[0].text

#     # 2. JSON 부분만 추출하여 파싱
#     if isinstance(response,str):
#         print('new_run is str')
#         new_runs = parse_llm_response(response)
#         new_runs = new_runs[0]
#     else:
#         new_runs = response
    
#     # print(new_runs)
#     # print(type(new_runs))
#     # 3. PPT 적용 로직 (기존과 동일)
#     old_range.Text = ""
#     current_range = old_range
    
#     for run_info in new_runs:
#         if not isinstance(run_info, dict):
#             continue
            
#         text_seg = run_info.get("Text", "")
#         if not text_seg:
#             continue

#         # InsertAfter는 삽입된 텍스트를 나타내는 새 TextRange 객체를 반환합니다.
#         new_run_range = current_range.InsertAfter(text_seg)
        
#         # 폰트 스타일 적용
#         font_info = run_info.get("Font", {})
#         f = new_run_range.Font
        
#         if font_info.get("Name"): f.Name = font_info["Name"]
#         if font_info.get("Size"): f.Size = font_info["Size"]
        
#         # win32com/VBA: True는 -1(msoTrue), False는 0(msoFalse)
#         f.Bold = -1 if font_info.get("Bold") else 0
#         f.Italic = -1 if font_info.get("Italic") else 0
#         f.Underline = -1 if font_info.get("Underline") else 0
        
#         # 색상 적용 (RGB)
#         color = font_info.get("Color")
#         if color and all(k in color for k in ("R", "G", "B")):
#             f.Color.RGB = color["R"] + (color["G"] << 8) + (color["B"] << 16)
        
#         # 다음 루프에서는 방금 삽입한 범위 뒤에 텍스트를 추가하도록 업데이트
#         current_range = new_run_range

#     return {
#         "operation": "replace_runs_llm",
#         "slide": slide_number,
#         "shape_id": shape_id,
#         "new_runs": new_runs,
#     }


def set_paragraph_alignment(prs, slide_number, shape_id, alignment="left", 
                           line_spacing=None, space_before=None, space_after=None):
    """
    Adjusts paragraph-level formatting.
    
    Args:
        alignment: 'left', 'center', 'right', 'justify', 'distribute'
        line_spacing: Line spacing multiplier (e.g., 1.5)
        space_before/after: Points before/after paragraph
    """
    shape = _find_shape_by_id(prs, slide_number, shape_id)
    if not shape.HasTextFrame:
        return f"Error: Shape {shape_id} cannot contain text."
    
    tr = shape.TextFrame.TextRange
    alignment_map = {
        'left': 1, 'center': 2, 'right': 3, 'justify': 4, 'distribute': 5
    }
    
    if alignment in alignment_map:
        tr.ParagraphFormat.Alignment = alignment_map[alignment]
    if line_spacing:
        tr.ParagraphFormat.LineRuleWithin = 0  # Multiple spacing
        tr.ParagraphFormat.SpaceWithin = line_spacing
    if space_before is not None:
        tr.ParagraphFormat.SpaceBefore = space_before
    if space_after is not None:
        tr.ParagraphFormat.SpaceAfter = space_after
    
    return f"Paragraph formatting applied to Shape {shape_id}."


def manage_bullet_points(prs, slide_number, shape_id, bullet_type="bullet", 
                     bullet_char=None, start_value=1):
    """
    Adds or modifies bullet/numbering format. 
    
    Args:
        bullet_type: 'bullet', 'number', 'none'
        bullet_char: Custom bullet character
        start_value: Starting number for numbered lists
    """
    shape = _find_shape_by_id(prs, slide_number, shape_id)
    if not shape.HasTextFrame:
        return f"Error: Shape {shape_id} cannot contain text."
    
    tr = shape.TextFrame.TextRange
    
    if bullet_type == "none":
        tr.ParagraphFormat.Bullet.Visible = False
    elif bullet_type == "bullet":
        tr.ParagraphFormat.Bullet.Visible = True
        tr.ParagraphFormat.Bullet.Type = 1  # ppBulletUnnumbered
        if bullet_char:
            tr.ParagraphFormat.Bullet.Character = ord(bullet_char)
    elif bullet_type == "number":
        tr.ParagraphFormat.Bullet.Visible = True
        tr.ParagraphFormat.Bullet.Type = 2  # ppBulletNumbered
        tr.ParagraphFormat.Bullet.StartValue = start_value
    
    return f"Bullet formatting applied to Shape {shape_id}."


# --- [B] Enhanced Text Content Editing ---

# def update_text(prs, slide_number, shape_id, new_text, append=False, 
#                 paragraph_index=None):
#     """
#     Updates text content with append option and paragraph targeting.
    
#     Args:
#         append: If True, adds to existing text instead of replacing
#         paragraph_index: Target specific paragraph (0-based)
#     """
#     shape = _find_shape_by_id(prs, slide_number, shape_id)
#     if not shape.HasTextFrame:
#         return f"Error: Shape {shape_id} does not support text editing."
    
#     if paragraph_index is not None:
#         para = shape.TextFrame.TextRange.Paragraphs(paragraph_index + 1)
#         if append:
#             para.Text = para.Text + new_text
#         else:
#             para.Text = new_text
#     else:
#         if append:
#             shape.TextFrame.TextRange.Text += new_text
#         else:
#             shape.TextFrame.TextRange.Text = new_text
    
#     return f"Successfully updated text for Shape {shape_id}."


def find_and_replace(prs, slide_number, shape_id, find_text, replace_text, 
                    match_case=False):
    """Finds and replaces text within a shape."""
    shape = _find_shape_by_id(prs, slide_number, shape_id)
    if not shape.HasTextFrame:
        return f"Error: Shape {shape_id} does not support text editing."
    
    current_text = shape.TextFrame.TextRange.Text
    
    if match_case:
        new_text = current_text.replace(find_text, replace_text)
    else:
        # Case-insensitive replace
        new_text = re.sub(re.escape(find_text), replace_text, current_text, flags=re.IGNORECASE)
    
    shape.TextFrame.TextRange.Text = new_text
    count = len(re.findall(re.escape(find_text), current_text, re.IGNORECASE if not match_case else 0))
    
    return f"Replaced {count} occurrence(s) in Shape {shape_id}."


# --- [C] Enhanced Layout / Geometry Editing ---

def adjust_layout(prs, slide_number, shape_id, left=None, top=None, 
                 width=None, height=None, rotation=None):
    """Adjusts position, size, and rotation of a shape."""
    shape = _find_shape_by_id(prs, slide_number, shape_id)
    
    if left is not None: shape.Left = left
    if top is not None: shape.Top = top
    if width is not None: shape.Width = width
    if height is not None: shape.Height = height
    if rotation is not None: shape.Rotation = rotation
    
    return f"Successfully adjusted layout for Shape {shape_id}."


def distribute_shapes(prs, slide_number, shape_ids, direction="horizontal", 
                     spacing=None):
    """
    Distributes multiple shapes evenly.
    
    Args:
        direction: 'horizontal' or 'vertical'
        spacing: Fixed spacing between shapes (if None, distribute evenly)
    """
    shapes = [_find_shape_by_id(prs, slide_number, sid) for sid in shape_ids]
    
    if len(shapes) < 2:
        return "Need at least 2 shapes to distribute."
    
    if direction == "horizontal":
        shapes.sort(key=lambda s: s.Left)
        if spacing:
            for i in range(1, len(shapes)):
                shapes[i].Left = shapes[i-1].Left + shapes[i-1].Width + spacing
        else:
            total_width = sum(s.Width for s in shapes)
            start = shapes[0].Left
            end = shapes[-1].Left + shapes[-1].Width
            available = end - start - total_width
            gap = available / (len(shapes) - 1) if len(shapes) > 1 else 0
            
            current_left = start
            for shape in shapes:
                shape.Left = current_left
                current_left += shape.Width + gap
    else:  # vertical
        shapes.sort(key=lambda s: s.Top)
        if spacing:
            for i in range(1, len(shapes)):
                shapes[i].Top = shapes[i-1].Top + shapes[i-1].Height + spacing
        else:
            total_height = sum(s.Height for s in shapes)
            start = shapes[0].Top
            end = shapes[-1].Top + shapes[-1].Height
            available = end - start - total_height
            gap = available / (len(shapes) - 1) if len(shapes) > 1 else 0
            
            current_top = start
            for shape in shapes:
                shape.Top = current_top
                current_top += shape.Height + gap
    
    return f"Distributed {len(shapes)} shapes {direction}ly."


def align_shapes(prs, slide_number, shape_ids, align_type="left"):
    """
    Aligns multiple shapes to each other.
    
    Args:
        align_type: 'left', 'right', 'top', 'bottom', 'center_h', 'center_v'
    """
    shapes = [_find_shape_by_id(prs, slide_number, sid) for sid in shape_ids]
    
    if len(shapes) < 2:
        return "Need at least 2 shapes to align."
    
    if align_type == "left":
        left_most = min(s.Left for s in shapes)
        for shape in shapes:
            shape.Left = left_most
    elif align_type == "right":
        right_most = max(s.Left + s.Width for s in shapes)
        for shape in shapes:
            shape.Left = right_most - shape.Width
    elif align_type == "top":
        top_most = min(s.Top for s in shapes)
        for shape in shapes:
            shape.Top = top_most
    elif align_type == "bottom":
        bottom_most = max(s.Top + s.Height for s in shapes)
        for shape in shapes:
            shape.Top = bottom_most - shape.Height
    elif align_type == "center_h":
        avg_center = sum(s.Left + s.Width / 2 for s in shapes) / len(shapes)
        for shape in shapes:
            shape.Left = avg_center - shape.Width / 2
    elif align_type == "center_v":
        avg_center = sum(s.Top + s.Height / 2 for s in shapes) / len(shapes)
        for shape in shapes:
            shape.Top = avg_center - shape.Height / 2
    
    return f"Aligned {len(shapes)} shapes by {align_type}."


# --- [D] Enhanced Object Lifecycle ---

def manage_object(prs, slide_number, action, shape_id=None, shape_type=1, 
                 left=100, top=100, width=100, height=100, text=None):
    """Creates, deletes, or duplicates shapes."""
    slide = prs.Slides(slide_number)
    
    if action == "add":
        new_shape = slide.Shapes.AddShape(shape_type, left, top, width, height)
        if text and new_shape.HasTextFrame:
            new_shape.TextFrame.TextRange.Text = text
        return f"Shape created successfully (ID: {new_shape.Id})."
    
    elif action == "delete" and shape_id:
        _find_shape_by_id(prs, slide_number, shape_id).Delete()
        return f"Shape {shape_id} deleted successfully."
    
    elif action == "duplicate" and shape_id:
        original = _find_shape_by_id(prs, slide_number, shape_id)
        duplicate = original.Duplicate()
        duplicate.Left += 20  # Offset slightly
        duplicate.Top += 20
        return f"Shape {shape_id} duplicated (New ID: {duplicate.Id})."
    
    return "Invalid action or missing shape_id."


def add_textbox(prs, slide_number, left, top, width, height, text=""):
    """Creates a new textbox with specified text."""
    slide = prs.Slides(slide_number)
    textbox = slide.Shapes.AddTextbox(1, left, top, width, height)  # msoTextOrientationHorizontal
    if text:
        textbox.TextFrame.TextRange.Text = text
    return f"Textbox created (ID: {textbox.Id})."


def add_image(prs, slide_number, image_path, left, top, width=None, height=None):
    """Inserts an image onto the slide."""
    slide = prs.Slides(slide_number)
    
    if width and height:
        picture = slide.Shapes.AddPicture(image_path, False, True, left, top, width, height)
    else:
        picture = slide.Shapes.AddPicture(image_path, False, True, left, top)
    
    return f"Image inserted (ID: {picture.Id})."


def group_shapes(prs, slide_number, shape_ids):
    """Groups multiple shapes together."""
    slide = prs.Slides(slide_number)
    shapes = [_find_shape_by_id(prs, slide_number, sid) for sid in shape_ids]
    
    # Create shape range
    shape_range = slide.Shapes.Range([s.Id for s in shapes])
    grouped = shape_range.Group()
    
    return f"Grouped {len(shape_ids)} shapes (Group ID: {grouped.Id})."


def ungroup_shapes(prs, slide_number, group_id):
    """Ungroups a grouped shape."""
    group = _find_shape_by_id(prs, slide_number, group_id)
    ungrouped = group.Ungroup()
    
    return f"Ungrouped shape {group_id} into {ungrouped.Count} shapes."


# --- [E] Enhanced Visual Style / Theme ---

def apply_visual_style(prs, slide_number, shape_id, bg_color_hex=None, 
                      line_color_hex=None, line_weight=None, line_style=None,
                      transparency=None, shadow=None):
    """
    Sets comprehensive visual styles.
    
    Args:
        line_style: 'solid', 'dash', 'dot', 'dash_dot'
        transparency: 0-1 (0=opaque, 1=transparent)
        shadow: True/False to enable/disable shadow
    """
    shape = _find_shape_by_id(prs, slide_number, shape_id)
    results = []

    if bg_color_hex:
        shape.Fill.Visible = True 
        shape.Fill.ForeColor.RGB = _hex_to_rgb_int(bg_color_hex)
        results.append(f"background({bg_color_hex})")

    if transparency is not None:
        shape.Fill.Transparency = transparency
        results.append(f"transparency({transparency})")

    if line_color_hex:
        shape.Line.Visible = True 
        shape.Line.ForeColor.RGB = _hex_to_rgb_int(line_color_hex)
        results.append(f"line color({line_color_hex})")

    if line_weight is not None:
        shape.Line.Visible = True
        shape.Line.Weight = line_weight
        results.append(f"line weight({line_weight}pt)")
    
    if line_style:
        shape.Line.Visible = True
        style_map = {'solid': 1, 'dash': 2, 'dot': 3, 'dash_dot': 4}
        if line_style in style_map:
            shape.Line.DashStyle = style_map[line_style]
            results.append(f"line style({line_style})")
    
    if shadow is not None:
        shape.Shadow.Visible = shadow
        results.append(f"shadow({'on' if shadow else 'off'})")

    return f"Shape {shape_id} visual style updated: " + ", ".join(results) if results else f"No changes applied to Shape {shape_id}."


def apply_gradient_fill(prs, slide_number, shape_id, color1_hex, color2_hex, 
                       gradient_type="linear", angle=0):
    """
    Applies gradient fill to a shape.
    
    Args:
        gradient_type: 'linear', 'radial', 'rectangular', 'path'
        angle: Gradient angle in degrees (for linear)
    """
    shape = _find_shape_by_id(prs, slide_number, shape_id)
    
    gradient_map = {'linear': 1, 'radial': 3, 'rectangular': 4, 'path': 5}
    
    shape.Fill.TwoColorGradient(gradient_map.get(gradient_type, 1), 1)
    shape.Fill.ForeColor.RGB = _hex_to_rgb_int(color1_hex)
    shape.Fill.BackColor.RGB = _hex_to_rgb_int(color2_hex)
    
    if gradient_type == "linear":
        # Set gradient angle
        shape.Fill.GradientAngle = angle
    
    return f"Gradient applied to Shape {shape_id}."


def set_shape_effect(prs, slide_number, shape_id, effect_type, **kwargs):
    """
    Applies special effects to shapes.
    
    Args:
        effect_type: 'glow', 'soft_edge', 'reflection', '3d'
        kwargs: Effect-specific parameters
    """
    shape = _find_shape_by_id(prs, slide_number, shape_id)
    
    if effect_type == "glow":
        color_hex = kwargs.get('color_hex', 'FFFF00')
        size = kwargs.get('size', 10)
        shape.Glow.Color.RGB = _hex_to_rgb_int(color_hex)
        shape.Glow.Radius = size
        return f"Glow effect applied to Shape {shape_id}."
    
    elif effect_type == "soft_edge":
        radius = kwargs.get('radius', 5)
        shape.SoftEdge.Radius = radius
        return f"Soft edge applied to Shape {shape_id}."
    
    elif effect_type == "reflection":
        shape.Reflection.Type = 1  # Enable reflection
        return f"Reflection applied to Shape {shape_id}."
    
    return f"Unknown effect type: {effect_type}"


# --- [F] Enhanced Consistency / Polishing ---

def align_to_object(prs, slide_number, target_id, base_id, side="right", margin=10):
    """Aligns the target shape relative to a base shape with custom margin."""
    target = _find_shape_by_id(prs, slide_number, target_id)
    base = _find_shape_by_id(prs, slide_number, base_id)
    
    if side == "right":
        target.Left = base.Left + base.Width + margin
        target.Top = base.Top
    elif side == "left":
        target.Left = base.Left - target.Width - margin
        target.Top = base.Top
    elif side == "bottom":
        target.Left = base.Left
        target.Top = base.Top + base.Height + margin
    elif side == "top":
        target.Left = base.Left
        target.Top = base.Top - target.Height - margin
    elif side == "center":
        target.Left = base.Left + (base.Width - target.Width) / 2
        target.Top = base.Top + (base.Height - target.Height) / 2
        
    return f"Aligned {target_id} to the {side} of {base_id}."


def match_formatting(prs, slide_number, source_id, target_ids):
    """Copies formatting from source shape to target shapes."""
    source = _find_shape_by_id(prs, slide_number, source_id)
    targets = [_find_shape_by_id(prs, slide_number, tid) for tid in target_ids]
    
    for target in targets:
        # Copy fill
        if source.Fill.Visible:
            target.Fill.ForeColor.RGB = source.Fill.ForeColor.RGB
            target.Fill.Transparency = source.Fill.Transparency
        
        # Copy line
        if source.Line.Visible:
            target.Line.ForeColor.RGB = source.Line.ForeColor.RGB
            target.Line.Weight = source.Line.Weight
        
        # Copy text format if applicable
        if source.HasTextFrame and target.HasTextFrame:
            src_tr = source.TextFrame.TextRange
            tgt_tr = target.TextFrame.TextRange
            tgt_tr.Font.Name = src_tr.Font.Name
            tgt_tr.Font.Size = src_tr.Font.Size
            tgt_tr.Font.Color.RGB = src_tr.Font.Color.RGB
            tgt_tr.Font.Bold = src_tr.Font.Bold
            tgt_tr.Font.Italic = src_tr.Font.Italic
    
    return f"Formatting copied from {source_id} to {len(target_ids)} shape(s)."


def set_z_order(prs, slide_number, shape_id, order="bring_to_front"):
    """
    Changes the z-order (layering) of a shape.
    
    Args:
        order: 'bring_to_front', 'send_to_back', 'bring_forward', 'send_backward'
    """
    shape = _find_shape_by_id(prs, slide_number, shape_id)
    
    if order == "bring_to_front":
        shape.ZOrder(0)  # msoBringToFront
    elif order == "send_to_back":
        shape.ZOrder(1)  # msoSendToBack
    elif order == "bring_forward":
        shape.ZOrder(2)  # msoBringForward
    elif order == "send_backward":
        shape.ZOrder(3)  # msoSendBackward
    
    return f"Z-order changed for Shape {shape_id}."


# --- [G] Slide Management ---

def add_slide(prs, layout_index=1, position=None):
    """
    Adds a new slide to the presentation.
    
    Args:
        layout_index: Layout to use (1-based)
        position: Where to insert (None = end)
    """
    layout = prs.SlideMaster.CustomLayouts(layout_index)
    
    if position:
        new_slide = prs.Slides.AddSlide(position, layout)
    else:
        new_slide = prs.Slides.Add(prs.Slides.Count + 1, layout)
    
    return f"Slide added at position {new_slide.SlideIndex}."


def delete_slide(prs, slide_number):
    """Deletes a specific slide."""
    prs.Slides(slide_number).Delete()
    return f"Slide {slide_number} deleted."


def duplicate_slide(prs, slide_number):
    """Duplicates a specific slide."""
    original = prs.Slides(slide_number)
    duplicate = original.Duplicate()
    return f"Slide {slide_number} duplicated to position {duplicate.SlideIndex}."


# --- [H] Table Operations ---

def add_table(prs, slide_number, rows, cols, left, top, width, height):
    """Creates a table on the slide."""
    slide = prs.Slides(slide_number)
    table = slide.Shapes.AddTable(rows, cols, left, top, width, height)
    return f"Table created (ID: {table.Id}, {rows}x{cols})."


def update_table_cell(prs, slide_number, table_id, row, col, text, 
                     font_size=None, color_hex=None, bg_color_hex=None):
    """Updates content and style of a specific table cell."""
    table_shape = _find_shape_by_id(prs, slide_number, table_id)
    
    if not table_shape.HasTable:
        return f"Error: Shape {table_id} is not a table."
    
    cell = table_shape.Table.Cell(row, col)
    cell.Shape.TextFrame.TextRange.Text = text
    
    if font_size:
        cell.Shape.TextFrame.TextRange.Font.Size = font_size
    if color_hex:
        cell.Shape.TextFrame.TextRange.Font.Color.RGB = _hex_to_rgb_int(color_hex)
    if bg_color_hex:
        cell.Shape.Fill.ForeColor.RGB = _hex_to_rgb_int(bg_color_hex)
    
    return f"Table cell ({row},{col}) updated."


# --- [I] Animation & Transition ---

def add_animation(prs, slide_number, shape_id, effect_type="appear", 
                 trigger="on_click", duration=1.0):
    """
    Adds animation to a shape.
    
    Args:
        effect_type: 'appear', 'fade', 'fly_in', 'zoom', etc.
        trigger: 'on_click', 'with_previous', 'after_previous'
        duration: Animation duration in seconds
    """
    slide = prs.Slides(slide_number)
    shape = _find_shape_by_id(prs, slide_number, shape_id)
    
    effect_map = {
        'appear': 1,  # msoAnimEffectAppear
        'fade': 10,   # msoAnimEffectFade
        'fly_in': 22, # msoAnimEffectFly
        'zoom': 88,   # msoAnimEffectZoom
    }
    
    effect = slide.TimeLine.MainSequence.AddEffect(
        shape, effect_map.get(effect_type, 1), trigger=1, index=-1
    )
    effect.Timing.Duration = duration
    
    return f"Animation '{effect_type}' added to Shape {shape_id}."


def set_slide_transition(prs, slide_number, transition_type="fade", 
                        duration=1.0, advance_on_time=None):
    """
    Sets slide transition effect.
    
    Args:
        transition_type: 'fade', 'push', 'wipe', 'split', etc.
        duration: Transition duration in seconds
        advance_on_time: Auto-advance after N seconds (None = manual)
    """
    slide = prs.Slides(slide_number)
    
    transition_map = {
        'fade': 1,
        'push': 13,
        'wipe': 15,
        'split': 14,
    }
    
    slide.SlideShowTransition.EntryEffect = transition_map.get(transition_type, 1)
    slide.SlideShowTransition.Duration = duration
    
    if advance_on_time:
        slide.SlideShowTransition.AdvanceOnTime = True
        slide.SlideShowTransition.AdvanceTime = advance_on_time
    else:
        slide.SlideShowTransition.AdvanceOnClick = True
    
    return f"Transition '{transition_type}' applied to slide {slide_number}."


FUNCTION_MAP = {
    # [A]  Text Style Editing 
    "set_text_style_preserve_runs":set_text_style_preserve_runs,
    # "insert_text_from_textbox":insert_text_from_textbox,
    # "delete_text_from_textbox":delete_text_from_textbox,
    # "replace_text_from_textbox":replace_text_from_textbox,
    "replace_shape_text":replace_shape_text,


    "set_paragraph_alignment": set_paragraph_alignment,
    "manage_bullet_points": manage_bullet_points,


    # "find_and_replace": find_and_replace,
    "adjust_layout": adjust_layout,
    "distribute_shapes": distribute_shapes,
    "align_shapes": align_shapes,
    "manage_object": manage_object,
    "add_textbox": add_textbox,
    "apply_visual_style": apply_visual_style,
    "apply_gradient_fill": apply_gradient_fill,
    "add_slide": add_slide,
    "delete_slide": delete_slide,
    "duplicate_slide": duplicate_slide
}












