from utils import parse_active_slide_objects, _call_gpt_api,  parse_llm_response
import os
import json
from datetime import datetime
import time
import win32com.client
# # Directory for extracted assets
# IMAGE_DIR = os.path.abspath("extracted_images")
# if not os.path.exists(IMAGE_DIR):
#     os.makedirs(IMAGE_DIR)

# # def _bgr_int_to_hex(rgb_int):
# #     """Converts Win32 RGB integer (BGR) to HEX string."""
# #     if rgb_int is None: return "000000"
# #     blue = (rgb_int >> 16) & 0xff
# #     green = (rgb_int >> 8) & 0xff
# #     red = rgb_int & 0xff
# #     return f"{red:02x}{green:02x}{blue:02x}".upper()

# def parse_ppt(prs):
#     """
#     Parses the PowerPoint presentation and extracts metadata, 
#     slides, shapes, text styles, and images into a structured dictionary.
    
#     :param prs: The win32com.client Presentation object.
#     :return: A dictionary containing the PPT structure.
#     """
#     data = {
#         "metadata": {
#             "title": prs.Name,
#             "width": prs.PageSetup.SlideWidth,
#             "height": prs.PageSetup.SlideHeight
#         },
#         "slides": []
#     }

#     for i, slide in enumerate(prs.Slides):
#         slide_index = i + 1
#         slide_info = {
#             "slide_index": slide_index,
#             "slide_id": slide.SlideID,
#             "layout": slide.Layout,
#             "objects": []
#         }

#         for shape in slide.Shapes:
#             obj = {
#                 "id": shape.Id,
#                 "name": shape.Name,
#                 "type": shape.Type, # MsoShapeType enum
#                 "rect": {
#                     "left": round(shape.Left, 2),
#                     "top": round(shape.Top, 2),
#                     "width": round(shape.Width, 2),
#                     "height": round(shape.Height, 2)
#                 },
#                 "rotation": shape.Rotation,
#                 "visible": bool(shape.Visible)
#             }

#             # 1. Extract Text Data
#             if shape.HasTextFrame:
#                 if shape.TextFrame.HasText:
#                     obj["content_type"] = "text"
#                     obj["text_content"] = shape.TextFrame.TextRange.Text.replace('\r', '\n')
#                     obj["paragraphs"] = []
                    
#                     text_range = shape.TextFrame.TextRange
#                     for p in range(1, text_range.Paragraphs().Count + 1):
#                         para = text_range.Paragraphs(p)
#                         p_data = {
#                             "text": para.Text.replace('\r', ''),
#                             "alignment": para.ParagraphFormat.Alignment,
#                             "runs": []
#                         }
#                         # Extracting specific styles from runs (for precision)
#                         for r in range(1, para.Runs().Count + 1):
#                             run = para.Runs(r)
#                             p_data["runs"].append({
#                                 "text": run.Text.replace('\r', ''),
#                                 "style": {
#                                     "font_name": run.Font.Name,
#                                     "size": run.Font.Size,
#                                     "bold": bool(run.Font.Bold),
#                                     "italic": bool(run.Font.Italic),
#                                     # "color": _bgr_int_to_hex(run.Font.Color.RGB)
#                                     "color": run.Font.Color.RGB
#                                 }
#                             })
#                         obj["paragraphs"].append(p_data)

#             # 2. Extract Table Data (Added Improvement)
#             elif shape.HasTable:
#                 obj["content_type"] = "table"
#                 obj["table_data"] = []
#                 table = shape.Table
#                 for r in range(1, table.Rows.Count + 1):
#                     row_data = []
#                     for c in range(1, table.Columns.Count + 1):
#                         cell_text = table.Cell(r, c).Shape.TextFrame.TextRange.Text
#                         row_data.append(cell_text.replace('\r', ''))
#                     obj["table_data"].append(row_data)

#             # 3. Extract Image Data
#             elif shape.Type == 13: # 13 = msoPicture
#                 obj["content_type"] = "image"
#                 img_filename = f"slide_{slide_index}_shape_{shape.Id}.png"
#                 img_path = os.path.join(IMAGE_DIR, img_filename)
                
#                 # Export shape as image if not already extracted
#                 try:
#                     shape.Export(img_path, 2) # 2 = ppShapeFormatPNG
#                     obj["image_path"] = img_path
#                 except Exception as e:
#                     obj["image_error"] = str(e)

#             slide_info["objects"].append(obj)
        
#         data["slides"].append(slide_info)

#     return data



class Parser:
    def __init__(self, json_data=None, baseline=False):
        """
        Args:
            json_data (dict, optional): JSON instructions with 'tasks'.
            baseline (bool): If True, ignore json_data and parse all slides.
        """
        self.json_data = json_data or {}
        self.tasks = self.json_data.get("tasks", [])
        self.baseline = baseline

    def process(self):
        """
        Process slide parsing based on baseline flag.

        Returns:
            dict: Parsed results. If baseline=True, returns dict with all slides under 'tasks'.
                  Otherwise, returns the original json_data updated with 'contents'.
        """
        if self.baseline:
            # 전체 슬라이드 수 가져오기
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            presentation = ppt_app.ActivePresentation
            total_slides = presentation.Slides.Count

            # 모든 슬라이드에 대해 parsing 수행하여 새로운 tasks 리스트 생성
            all_tasks = []
            for page_num in range(1, total_slides + 1):
                slide_contents = parse_active_slide_objects(page_num)
                
                all_tasks.append({
                        "page number": page_num, 
                        "contents": slide_contents
                    })

            return all_tasks

        # baseline=False: 지정된 task들만 parsing
        for task in self.tasks:
            page_number = task.get("page number")
            if page_number:
                slide_contents = parse_active_slide_objects(page_number)
                task["contents"] = slide_contents

        return self.json_data
import re
import json
import ast

def parse_llm_response_processor(response):
    """
    Robustly parse JSON or Python-like structures from an LLM response.
    Returns the loaded object (dict or list), or None if parsing fails.
    """
    if not response or not isinstance(response, str):
        return None

    # Remove markdown code fences
    response_clean = re.sub(r'```(?:json|python)?', '', response).strip()
    
    # Extract JSON or Python literal between the outermost { } or [ ]
    match = re.search(r'(\{[\s\S]*\}|\[[\s\S]*\])', response_clean)
    if not match:
        return None

    payload = match.group(1)
    
    # Remove trailing commas before } or ]
    payload = re.sub(r',(\s*[\}\]])', r'\1', payload)
    
    # Try JSON parsing first
    try:
        return json.loads(payload)
    except json.JSONDecodeError:
        # Convert Python style to JSON style for json.loads
        try:
            # Replace Python booleans/None with JSON equivalents
            json_payload = payload.replace('True', 'true').replace('False', 'false').replace('None', 'null')
            return json.loads(json_payload)
        except json.JSONDecodeError:
            # Fallback to Python literal eval
            try:
                # Replace JSON literals with Python equivalents
                python_payload = payload.replace('null', 'None').replace('true', 'True').replace('false', 'False')
                return ast.literal_eval(python_payload)
            except (ValueError, SyntaxError):
                return None
            