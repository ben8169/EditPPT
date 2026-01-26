# import win32com.client
# import json


# # --- Internal Helper Functions ---

# def _hex_to_rgb_int(hex_str):
#     """Converts HEX string (#FFFFFF or FFFFFF) to win32-compatible BGR integer."""
#     hex_str = hex_str.lstrip('#')
#     if len(hex_str) != 6:
#         raise ValueError("HEX code must be 6 characters long (e.g., FF0000).")
    
#     r = int(hex_str[0:2], 16)
#     g = int(hex_str[2:4], 16)
#     b = int(hex_str[4:6], 16)
#     # Win32com uses BGR structure: (b << 16) | (g << 8) | r
#     return (b << 16) | (g << 8) | r

# def find_shape_by_id(prs, slide_index, shape_id):
#     """Finds a specific Shape object by its unique ID on a given slide."""
#     try:
#         # Note: slide_index usually starts from 1 in PPT API
#         slide = prs.Slides(slide_index)
#         for shape in slide.Shapes:
#             if shape.Id == shape_id:
#                 return shape
#     except Exception as e:
#         raise ValueError(f"Error accessing slide {slide_index}: {e}")
        
#     raise ValueError(f"Shape with ID {shape_id} not found on slide {slide_index}.")


# # --- [A] Text Style Editing ---

# def set_text_style(prs, slide_index, shape_id, font_size=None, color_hex=None, bold=None):
#     """Modifies the font size, color, and weight of the text in a shape."""
#     shape = find_shape_by_id(prs, slide_index, shape_id)
#     if not shape.HasTextFrame:
#         return f"Error: Shape {shape_id} cannot contain text."
    
#     tr = shape.TextFrame.TextRange
#     if font_size: tr.Font.Size = font_size
#     if bold is not None: tr.Font.Bold = bold
#     if color_hex:
#         tr.Font.Color.RGB = _hex_to_rgb_int(color_hex)
    
#     return f"Successfully updated style for Shape {shape_id}."


# # --- [B] Text Content Editing ---

def update_text(prs, slide_index, shape_id, new_text):
    """Updates the actual text content inside a shape."""
    shape = find_shape_by_id(prs, slide_index, shape_id)
    if not shape.HasTextFrame:
        return f"Error: Shape {shape_id} does not support text editing."
    
    shape.TextFrame.TextRange.Text = new_text
    return f"Successfully updated text for Shape {shape_id}."


# # --- [C] Layout / Geometry Editing ---

# def adjust_layout(prs, slide_index, shape_id, left=None, top=None, width=None, height=None):
#     """Adjusts the position (left, top) and size (width, height) of a shape."""
#     shape = find_shape_by_id(prs, slide_index, shape_id)
#     if left is not None: shape.Left = left
#     if top is not None: shape.Top = top
#     if width is not None: shape.Width = width
#     if height is not None: shape.Height = height
#     return f"Successfully adjusted layout for Shape {shape_id}."


# # --- [D] Object Lifecycle ---

# def manage_object(prs, slide_index, action, shape_id=None, shape_type=1, left=100, top=100, width=100, height=100):
#     """Adds a new shape or deletes an existing one."""
#     slide = prs.Slides(slide_index)
#     if action == "add":
#         new_shape = slide.Shapes.AddShape(shape_type, left, top, width, height)
#         return f"Shape created successfully (ID: {new_shape.Id})."
#     elif action == "delete" and shape_id:
#         find_shape_by_id(prs, slide_index, shape_id).Delete()
#         return f"Shape {shape_id} deleted successfully."
#     return "Invalid action or missing shape_id."


# # --- [E] Visual Style / Theme ---

# def apply_visual_style(prs, slide_index, shape_id, bg_color_hex=None, line_color_hex=None, line_weight=None):
#     """Sets the background fill color, border color, and border weight of a shape."""
#     shape = find_shape_by_id(prs, slide_index, shape_id)
#     results = []

#     if bg_color_hex:
#         shape.Fill.Visible = True 
#         shape.Fill.ForeColor.RGB = _hex_to_rgb_int(bg_color_hex)
#         results.append(f"background({bg_color_hex})")

#     if line_color_hex:
#         shape.Line.Visible = True 
#         shape.Line.ForeColor.RGB = _hex_to_rgb_int(line_color_hex)
#         results.append(f"line color({line_color_hex})")

#     if line_weight is not None:
#         shape.Line.Visible = True
#         shape.Line.Weight = line_weight
#         results.append(f"line weight({line_weight}pt)")

#     if not results:
#         return f"No changes applied to Shape {shape_id}."
    
#     return f"Shape {shape_id} visual style updated: " + ", ".join(results)


# # --- [F] Consistency / Polishing ---

# def align_to_object(prs, slide_index, target_id, base_id, side="right"):
#     """Aligns the target shape relative to a base shape."""
#     target = find_shape_by_id(prs, slide_index, target_id)
#     base = find_shape_by_id(prs, slide_index, base_id)
    
#     margin = 10 # Default spacing
    
#     if side == "right":
#         target.Left = base.Left + base.Width + margin
#         target.Top = base.Top
#     elif side == "left":
#         target.Left = base.Left - target.Width - margin
#         target.Top = base.Top
#     elif side == "bottom":
#         target.Left = base.Left
#         target.Top = base.Top + base.Height + margin
#     elif side == "top":
#         target.Left = base.Left
#         target.Top = base.Top - target.Height - margin
#     elif side == "center":
#         target.Left = base.Left + (base.Width - target.Width) / 2
#         target.Top = base.Top + (base.Height - target.Height) / 2
        
#     return f"Aligned {target_id} to the {side} of {base_id}."

# def sum_numbers(a: float, b: float) -> dict:
#     """Calculates the sum of two numbers."""
#     return {"result": a + b}


# # --- Tool Schema ---

# TOOLS_SCHEMA = [
#     {
#         "type": "function",
#         "function": {
#             "name": "set_text_style",
#             "description": "Precisely adjusts text styles including font size, color, and bold properties.",
#             "parameters": {
#                 "type": "object",
#                 "properties": {
#                     "slide_index": {"type": "integer", "description": "1-based index of the slide."},
#                     "shape_id": {"type": "integer", "description": "The unique ID of the shape."},
#                     "font_size": {"type": "number", "description": "Font size in points (pt)."},
#                     "color_hex": {"type": "string", "description": "Hex color code (e.g., FF0000)."},
#                     "bold": {"type": "boolean", "description": "Whether to make text bold."}
#                 },
#                 "required": ["slide_index", "shape_id"]
#             }
#         }
#     },
#     {
#         "type": "function",
#         "function": {
#             "name": "update_text",
#             "description": "Updates the text content inside a specific shape.",
#             "parameters": {
#                 "type": "object",
#                 "properties": {
#                     "slide_index": {"type": "integer"},
#                     "shape_id": {"type": "integer"},
#                     "new_text": {"type": "string", "description": "The new text string to insert."}
#                 },
#                 "required": ["slide_index", "shape_id", "new_text"]
#             }
#         }
#     },
#     {
#         "type": "function",
#         "function": {
#             "name": "adjust_layout",
#             "description": "Adjusts the position and dimensions of a shape.",
#             "parameters": {
#                 "type": "object",
#                 "properties": {
#                     "slide_index": {"type": "integer"},
#                     "shape_id": {"type": "integer"},
#                     "left": {"type": "number"},
#                     "top": {"type": "number"},
#                     "width": {"type": "number"},
#                     "height": {"type": "number"}
#                 },
#                 "required": ["slide_index", "shape_id"]
#             }
#         }
#     },
#     {
#         "type": "function",
#         "function": {
#             "name": "manage_object",
#             "description": "Creates or deletes shapes. (Heart: 21, Star: 9, Rectangle: 1, Oval: 9)",
#             "parameters": {
#                 "type": "object",
#                 "properties": {
#                     "slide_index": {"type": "integer"},
#                     "action": {"type": "string", "enum": ["add", "delete"]},
#                     "shape_id": {"type": "integer", "description": "Required only for delete action."},
#                     "shape_type": {"type": "integer", "default": 1},
#                     "left": {"type": "number", "default": 100},
#                     "top": {"type": "number", "default": 100},
#                     "width": {"type": "number", "default": 100},
#                     "height": {"type": "number", "default": 100}
#                 },
#                 "required": ["slide_index", "action"]
#             }
#         }
#     },
#     {
#         "type": "function",
#         "function": {
#             "name": "apply_visual_style",
#             "description": "Sets the background color, line color, and line weight of a shape.",
#             "parameters": {
#                 "type": "object",
#                 "properties": {
#                     "slide_index": {"type": "integer"},
#                     "shape_id": {"type": "integer"},
#                     "bg_color_hex": {"type": "string", "description": "Background HEX code."},
#                     "line_color_hex": {"type": "string", "description": "Line/Border HEX code."},
#                     "line_weight": {"type": "number", "description": "Line thickness in pt."}
#                 },
#                 "required": ["slide_index", "shape_id"]
#             }
#         }
#     },
#     {
#         "type": "function",
#         "function": {
#             "name": "align_to_object",
#             "description": "Aligns a target shape relative to a base shape.",
#             "parameters": {
#                 "type": "object",
#                 "properties": {
#                     "slide_index": {"type": "integer"},
#                     "target_id": {"type": "integer"},
#                     "base_id": {"type": "integer"},
#                     "side": {"type": "string", "enum": ["left", "right", "top", "bottom", "center"]}
#                 },
#                 "required": ["slide_index", "target_id", "base_id", "side"]
#             }
#         }
#     },
#     {
#         "type": "function",
#         "function": {
#             "name": "sum_numbers",
#             "description": "Calculates the sum of two numbers.",
#             "parameters": {
#                 "type": "object",
#                 "properties": {
#                     "a": {"type": "number"},
#                     "b": {"type": "number"}
#                 },
#                 "required": ["a", "b"]
#             }
#         }
#     }
# ]

# # --- Function Map ---

# FUNCTION_MAP = {
#     "set_text_style": set_text_style,
#     "update_text": update_text,
#     "adjust_layout": adjust_layout,
#     "manage_object": manage_object,
#     "apply_visual_style": apply_visual_style,
#     "align_to_object": align_to_object,
#     "sum_numbers": sum_numbers,
# }

from win32com.client import constants

# =====================================================
# Internal Helpers (NOT tools)
# =====================================================

def _hex_to_rgb_int(hex_str):
    """Converts HEX string (#RRGGBB) to win32 BGR integer."""
    hex_str = hex_str.lstrip("#")
    if len(hex_str) != 6:
        raise ValueError("HEX must be 6 characters.")
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    return (b << 16) | (g << 8) | r


def find_shape_by_id(prs, slide_index, shape_id):
    slide = prs.Slides(slide_index)
    for shape in slide.Shapes:
        if shape.Id == shape_id:
            return shape
    raise ValueError(f"Shape with ID {shape_id} not found on slide {slide_index}.")


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
    slide_index,
    shape_id,
    container="shape",
    row_index=None,
    col_index=None
):
    shape = find_shape_by_id(prs, slide_index, shape_id)

    if container == "shape":
        return _get_text_runs_from_shape(shape)

    if container == "table_cell":
        if row_index is None or col_index is None:
            raise ValueError("row_index and col_index required for table_cell.")
        return _get_text_runs_from_table_cell(shape, row_index, col_index)

    raise ValueError(f"Unknown container type: {container}")


# =====================================================
# [A] Text Style Editing (Run ONLY)
# =====================================================

def set_text_run_style(
    prs,
    slide_index,
    shape_id,
    run_index,
    font_size=None,
    color_hex=None,
    bold=None,
    italic=None,
    container="shape",
    row_index=None,
    col_index=None
):
    """Modify font style of a single text run."""
    try:
        runs = _resolve_runs(prs, slide_index, shape_id, container, row_index, col_index)
        run = runs(run_index)
        font = run.Font

        if font_size is not None:
            font.Size = font_size
        if bold is not None:
            font.Bold = bold
        if italic is not None:
            font.Italic = italic
        if color_hex:
            font.Color.RGB = _hex_to_rgb_int(color_hex)

        return f"Updated style of run {run_index} in Shape {shape_id}."
    except Exception as e:
        return f"Error updating run style: {e}"


# =====================================================
# [B] Text Content Editing (Run ONLY)
# =====================================================

def update_text_run(
    prs,
    slide_index,
    shape_id,
    run_index,
    new_text,
    container="shape",
    row_index=None,
    col_index=None
):
    """Update text of a single run."""
    try:
        runs = _resolve_runs(prs, slide_index, shape_id, container, row_index, col_index)
        runs(run_index).Text = new_text
        return f"Updated text of run {run_index} in Shape {shape_id}."
    except Exception as e:
        return f"Error updating run text: {e}"


def delete_text_run(
    prs,
    slide_index,
    shape_id,
    run_index,
    container="shape",
    row_index=None,
    col_index=None
):
    """Delete a single run."""
    try:
        runs = _resolve_runs(prs, slide_index, shape_id, container, row_index, col_index)
        runs(run_index).Delete()
        return f"Deleted run {run_index} from Shape {shape_id}."
    except Exception as e:
        return f"Error deleting run: {e}"


# =====================================================
# [C] Layout / Geometry
# =====================================================

def set_object_position(prs, slide_index, shape_id, left, top):
    """Set absolute position of a shape."""
    shape = find_shape_by_id(prs, slide_index, shape_id)
    shape.Left = left
    shape.Top = top
    return f"Moved Shape {shape_id} to ({left}, {top})."


def set_object_size(prs, slide_index, shape_id, width, height):
    """Set absolute size of a shape."""
    shape = find_shape_by_id(prs, slide_index, shape_id)
    shape.Width = width
    shape.Height = height
    return f"Resized Shape {shape_id} to {width}x{height}."


# =====================================================
# [D] Object Lifecycle
# =====================================================

def create_shape(prs, slide_index, shape_type, left, top, width, height):
    """Create a new shape."""
    slide = prs.Slides(slide_index)
    new_shape = slide.Shapes.AddShape(shape_type, left, top, width, height)
    return f"Created Shape (ID: {new_shape.Id})."


def delete_shape(prs, slide_index, shape_id):
    """Delete a shape."""
    shape = find_shape_by_id(prs, slide_index, shape_id)
    shape.Delete()
    return f"Deleted Shape {shape_id}."


# =====================================================
# [E] Visual Style
# =====================================================

def set_shape_fill_color(prs, slide_index, shape_id, color_hex):
    """Set solid fill color of a shape."""
    shape = find_shape_by_id(prs, slide_index, shape_id)
    shape.Fill.Visible = True
    shape.Fill.ForeColor.RGB = _hex_to_rgb_int(color_hex)
    return f"Updated fill color for Shape {shape_id}."


def set_shape_outline(prs, slide_index, shape_id, color_hex=None, weight=None):
    """Set outline color and/or weight."""
    shape = find_shape_by_id(prs, slide_index, shape_id)
    shape.Line.Visible = True

    if color_hex:
        shape.Line.ForeColor.RGB = _hex_to_rgb_int(color_hex)
    if weight is not None:
        shape.Line.Weight = weight

    return f"Updated outline for Shape {shape_id}."


# =====================================================
# [F] Consistency / Polishing
# =====================================================

def align_shape_to_shape(prs, slide_index, target_shape_id, base_shape_id, axis):
    """Align one shape relative to another."""
    slide = prs.Slides(slide_index)
    t = find_shape_by_id(prs, slide_index, target_shape_id)
    b = find_shape_by_id(prs, slide_index, base_shape_id)

    if axis == "left":
        t.Left = b.Left
    elif axis == "right":
        t.Left = b.Left + b.Width - t.Width
    elif axis == "top":
        t.Top = b.Top
    elif axis == "bottom":
        t.Top = b.Top + b.Height - t.Height
    elif axis == "center_x":
        t.Left = b.Left + (b.Width - t.Width) / 2
    elif axis == "center_y":
        t.Top = b.Top + (b.Height - t.Height) / 2
    else:
        return f"Unknown alignment axis: {axis}"

    return f"Aligned Shape {target_shape_id} to Shape {base_shape_id} ({axis})."


# =====================================================

# =====================================================
# Tool Schema
# =====================================================

TOOLS_SCHEMA = [

    # =================================================
    # [A] Text Style Editing (Run ONLY)
    # =================================================

    {
        "type": "function",
        "function": {
            "name": "set_text_run_style",
            "description": "Modify font style of a specific text run within a shape or table cell.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {
                        "type": "integer",
                        "description": "1-based index of the slide."
                    },
                    "shape_id": {
                        "type": "integer",
                        "description": "Unique ID of the shape containing the text."
                    },
                    "run_index": {
                        "type": "integer",
                        "description": "1-based index of the text run."
                    },
                    "font_size": {
                        "type": "number",
                        "description": "Font size in points."
                    },
                    "color_hex": {
                        "type": "string",
                        "description": "Font color as HEX string (e.g., FF0000)."
                    },
                    "bold": {
                        "type": "boolean",
                        "description": "Whether the text is bold."
                    },
                    "italic": {
                        "type": "boolean",
                        "description": "Whether the text is italic."
                    },
                    "container": {
                        "type": "string",
                        "enum": ["shape", "table_cell"],
                        "description": "Text container type.",
                        "default": "shape"
                    },
                    "row_index": {
                        "type": "integer",
                        "description": "Table row index (required if container is table_cell)."
                    },
                    "col_index": {
                        "type": "integer",
                        "description": "Table column index (required if container is table_cell)."
                    }
                },
                "required": ["slide_index", "shape_id", "run_index"]
            }
        }
    },

    # =================================================
    # [B] Text Content Editing (Run ONLY)
    # =================================================

    {
        "type": "function",
        "function": {
            "name": "update_text_run",
            "description": "Update the text content of a specific run.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "shape_id": {"type": "integer"},
                    "run_index": {"type": "integer"},
                    "new_text": {
                        "type": "string",
                        "description": "New text content for the run."
                    },
                    "container": {
                        "type": "string",
                        "enum": ["shape", "table_cell"],
                        "default": "shape"
                    },
                    "row_index": {"type": "integer"},
                    "col_index": {"type": "integer"}
                },
                "required": ["slide_index", "shape_id", "run_index", "new_text"]
            }
        }
    },

    {
        "type": "function",
        "function": {
            "name": "delete_text_run",
            "description": "Delete a specific text run.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "shape_id": {"type": "integer"},
                    "run_index": {"type": "integer"},
                    "container": {
                        "type": "string",
                        "enum": ["shape", "table_cell"],
                        "default": "shape"
                    },
                    "row_index": {"type": "integer"},
                    "col_index": {"type": "integer"}
                },
                "required": ["slide_index", "shape_id", "run_index"]
            }
        }
    },

    # =================================================
    # [C] Layout / Geometry
    # =================================================

    {
        "type": "function",
        "function": {
            "name": "set_object_position",
            "description": "Set absolute position of a shape.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "shape_id": {"type": "integer"},
                    "left": {"type": "number"},
                    "top": {"type": "number"}
                },
                "required": ["slide_index", "shape_id", "left", "top"]
            }
        }
    },

    {
        "type": "function",
        "function": {
            "name": "set_object_size",
            "description": "Set absolute size of a shape.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "shape_id": {"type": "integer"},
                    "width": {"type": "number"},
                    "height": {"type": "number"}
                },
                "required": ["slide_index", "shape_id", "width", "height"]
            }
        }
    },

    # =================================================
    # [D] Object Lifecycle
    # =================================================

    {
        "type": "function",
        "function": {
            "name": "create_shape",
            "description": "Create a new shape on a slide.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "shape_type": {
                        "type": "integer",
                        "description": "PowerPoint shape type constant."
                    },
                    "left": {"type": "number"},
                    "top": {"type": "number"},
                    "width": {"type": "number"},
                    "height": {"type": "number"}
                },
                "required": ["slide_index", "shape_type", "left", "top", "width", "height"]
            }
        }
    },

    {
        "type": "function",
        "function": {
            "name": "delete_shape",
            "description": "Delete an existing shape.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "shape_id": {"type": "integer"}
                },
                "required": ["slide_index", "shape_id"]
            }
        }
    },

    # =================================================
    # [E] Visual Style
    # =================================================

    {
        "type": "function",
        "function": {
            "name": "set_shape_fill_color",
            "description": "Set solid fill color of a shape.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "shape_id": {"type": "integer"},
                    "color_hex": {
                        "type": "string",
                        "description": "Fill color HEX (e.g., 00FF00)."
                    }
                },
                "required": ["slide_index", "shape_id", "color_hex"]
            }
        }
    },

    {
        "type": "function",
        "function": {
            "name": "set_shape_outline",
            "description": "Set outline color and/or weight of a shape.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "shape_id": {"type": "integer"},
                    "color_hex": {"type": "string"},
                    "weight": {
                        "type": "number",
                        "description": "Line thickness in points."
                    }
                },
                "required": ["slide_index", "shape_id"]
            }
        }
    },

    # =================================================
    # [F] Consistency / Polishing
    # =================================================

    {
        "type": "function",
        "function": {
            "name": "align_shape_to_shape",
            "description": "Align one shape relative to another.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "target_shape_id": {"type": "integer"},
                    "base_shape_id": {"type": "integer"},
                    "axis": {
                        "type": "string",
                        "enum": [
                            "left",
                            "right",
                            "top",
                            "bottom",
                            "center_x",
                            "center_y"
                        ]
                    }
                },
                "required": [
                    "slide_index",
                    "target_shape_id",
                    "base_shape_id",
                    "axis"
                ]
            }
        }
    }
]

# =====================================================

# =====================================================
# Function Map
# =====================================================

FUNCTION_MAP = {
    # [A] Text Style Editing (Run ONLY)
    "set_text_run_style": set_text_run_style,

    # [B] Text Content Editing (Run ONLY)
    "update_text_run": update_text_run,
    "delete_text_run": delete_text_run,

    # [C] Layout / Geometry
    "set_object_position": set_object_position,
    "set_object_size": set_object_size,

    # [D] Object Lifecycle
    "create_shape": create_shape,
    "delete_shape": delete_shape,

    # [E] Visual Style
    "set_shape_fill_color": set_shape_fill_color,
    "set_shape_outline": set_shape_outline,

    # [F] Consistency / Polishing
    "align_shape_to_shape": align_shape_to_shape,
}









############# Legacy (26.01.21)#################

# def set_text_style(prs, slide_number, shape_id, font_size=None, color_hex=None, 
#                    bold=None, italic=None, underline=None, font_name=None,
#                    paragraph_index=None, char_start=None, char_end=None):
#     """
#     Modifies text styles with support for partial text selection.
    
#     Args:
#         paragraph_index: Specific paragraph to style (0-based, None = all)
#         char_start, char_end: Character range within text (None = all)
#     """
#     shape = find_shape_by_id(prs, slide_number, shape_id)
#     if not shape.HasTextFrame:
#         return f"Error: Shape {shape_id} cannot contain text."
    
#     # Select target text range
#     if paragraph_index is not None:
#         tr = shape.TextFrame.TextRange.Paragraphs(paragraph_index + 1)
#     elif char_start is not None and char_end is not None:
#         tr = shape.TextFrame.TextRange.Characters(char_start, char_end - char_start)
#     else:
#         tr = shape.TextFrame.TextRange
    
#     # Apply styles
#     if font_size: tr.Font.Size = font_size
#     if bold is not None: tr.Font.Bold = bold
#     if italic is not None: tr.Font.Italic = italic
#     if underline is not None: tr.Font.Underline = underline
#     if font_name: tr.Font.Name = font_name
#     if color_hex: tr.Font.Color.RGB = _hex_to_rgb_int(color_hex)
    
#     return f"Successfully updated style for Shape {shape_id}."



#####################################################################################
############# Legacy (26.01.24)#################
#####################################################################################

# def set_text_style_by_char_range(
#     prs,
#     slide_number: int,
#     shape_id: int,
#     char_start: int,
#     target_text: str,
#     *,
#     char_end: int = None,
#     container: str = "shape",
#     row_index: int = None,
#     col_index: int = None,
#     font_name: str = None,
#     font_size: Union[int, float] = None,
#     bold: Optional[bool] = None,
#     italic: Optional[bool] = None,
#     underline: Optional[bool] = None,
#     color_hex: str = None,
# ):
#     # 1. live text read
#     text, _ = _get_text_with_offsets(
#         prs,
#         slide_number,
#         shape_id,
#         container=container,
#         row_index=row_index,
#         col_index=col_index,
#     )

#     # 2. normalize range
#     start, end = _normalize_char_range(
#         text=text,
#         char_start=char_start,
#         target_text=target_text,
#         char_end=char_end,
#     )

#     # 3. resolve TextRange
#     shape = _find_shape_by_id(prs, slide_number, shape_id)

#     if container == "shape":
#         tr = shape.TextFrame.TextRange
#     else:
#         tr = shape.Table.Cell(row_index, col_index).Shape.TextFrame.TextRange

#     # PowerPoint is 1-based
#     target = tr.Characters(start + 1, end - start)
#     font = target.Font

#     # 4. apply styles
#     if font_name is not None:
#         font.Name = font_name
#     if font_size is not None:
#         font.Size = font_size
#     if bold is not None:
#         font.Bold = int(bold)
#     if italic is not None:
#         font.Italic = int(italic)
#     if underline is not None:
#         font.Underline = int(underline)
#     if color_hex is not None:
#         font.Color.RGB = _hex_to_rgb_int(color_hex)

#     return {
#         "applied_range": [start, end],
#         "text": target_text,
#         "shape_id": shape_id,
#         "slide": slide_number
#     }

    # {
    # "type": "function",
    # "name": "set_text_style_by_char_range",
    # "description": "Modify text style (font, size, color, emphasis) of a specific character range without changing text content.",
    # "parameters": {
    #     "type": "object",
    #     "properties": {
    #     "slide_number": {
    #         "type": "integer",
    #         "description": "1-based slide index"
    #     },
    #     "shape_id": {
    #         "type": "integer",
    #         "description": "PowerPoint Shape ID"
    #     },
    #     "char_start": {
    #         "type": "integer",
    #         "description": "0-based character start index (trusted)"
    #     },
    #     "target_text": {
    #         "type": "string",
    #         "description": "Exact text to be styled (used for range validation)"
    #     },
    #     "char_end": {
    #         "type": "integer",
    #         "description": "Optional end index (used only for fallback validation)"
    #     },

    #     "container": {
    #         "type": "string",
    #         "enum": ["shape", "table_cell"],
    #         "default": "shape"
    #     },
    #     "row_index": {
    #         "type": "integer",
    #         "description": "1-based row index (required if container=table_cell)"
    #     },
    #     "col_index": {
    #         "type": "integer",
    #         "description": "1-based column index (required if container=table_cell)"
    #     },

    #     "font_name": { "type": "string" },
    #     "font_size": { "type": "number" },
    #     "bold": { "type": "boolean" },
    #     "italic": { "type": "boolean" },
    #     "underline": { "type": "boolean" },
    #     "color_hex": {
    #         "type": "string",
    #         "description": "HEX color string, e.g. #FF0000"
    #     }
    #     },
    #     "required": [
    #     "slide_number",
    #     "shape_id",
    #     "char_start",
    #     "target_text"
    #     ]
    # }
    # },

#####################################################################################
# def edit_text_content_by_char_range(
#     prs,
#     slide_number: int,
#     shape_id: int,
#     char_start: int,
#     target_text: str,
#     *,
#     operation: str,                 # "insert" | "replace" | "delete"
#     new_text: str = "",
#     char_end: int = None,
#     container: str = "shape",
#     row_index: int = None,
#     col_index: int = None,
# ):
#     """
#     Edit text content ONLY (no style changes).
#     """

#     text, _ = _get_text_with_offsets(
#         prs,
#         slide_number,
#         shape_id,
#         container=container,
#         row_index=row_index,
#         col_index=col_index,
#     )

#     start, end = _normalize_char_range(
#         text=text,
#         char_start=char_start,
#         target_text=target_text,
#         char_end=char_end,
#     )

#     shape = _find_shape_by_id(prs, slide_number, shape_id)

#     if container == "shape":
#         tr = shape.TextFrame.TextRange
#     else:
#         tr = shape.Table.Cell(row_index, col_index).Shape.TextFrame.TextRange

#     # PowerPoint is 1-based
#     start_pos = start + 1
#     length = end - start

#     target_range = tr.Characters(start_pos, length)

#     # 4️⃣ apply operation (STYLE PRESERVED)
#     if operation == "insert":
#         # insert AFTER the target range
#         target_range.InsertAfter(new_text)

#     elif operation == "delete":
#         target_range.Delete()

#     elif operation == "replace":
#         # ⭐ 핵심: anchor 먼저 확보
#         anchor = tr.Characters(start_pos, 0)

#         # 1) 기존 텍스트 삭제
#         target_range.Delete()

#         # 2) anchor 위치에 삽입
#         anchor.InsertAfter(new_text)
        

#     else:
#         raise ValueError(f"Unknown operation: {operation}")

#     return {
#         "operation": operation,
#         "range": [start, end],
#         "original_text": target_text,
#         "new_text": new_text,
#         "shape_id": shape_id,
#         "slide": slide_number
#     }


# {
#     "type": "function",
#     "name": "edit_text_content_by_char_range",
#     "description": "Edit text content only (insert, replace, delete) while preserving existing text styles.",
#     "parameters": {
#         "type": "object",
#         "properties": {
#         "slide_number": {
#             "type": "integer",
#             "description": "1-based slide index"
#         },
#         "shape_id": {
#             "type": "integer",
#             "description": "PowerPoint Shape ID"
#         },
#         "char_start": {
#             "type": "integer",
#             "description": "0-based character start index (trusted)"
#         },
#         "target_text": {
#             "type": "string",
#             "description": "Exact text to be edited (used for range validation)"
#         },
#         "char_end": {
#             "type": "integer",
#             "description": "Optional end index (used only for fallback validation)"
#         },

#         "operation": {
#             "type": "string",
#             "enum": ["insert", "replace", "delete"],
#             "description": "Type of text edit operation"
#         },
#         "new_text": {
#             "type": "string",
#             "description": "Text to insert or replace with (ignored for delete)",
#             "default": ""
#         },

#         "container": {
#             "type": "string",
#             "enum": ["shape", "table_cell"],
#             "default": "shape"
#         },
#         "row_index": {
#             "type": "integer",
#             "description": "1-based row index (required if container=table_cell)"
#         },
#         "col_index": {
#             "type": "integer",
#             "description": "1-based column index (required if container=table_cell)"
#         }
#         },
#         "required": [
#         "slide_number",
#         "shape_id",
#         "char_start",
#         "target_text",
#         "operation"
#         ]
#     }
#     }
#####################################################################################



#####################################################################################
############# Legacy (26.01.26)#################
#####################################################################################

# def _resolve_runs(
#     prs,
#     slide_number,
#     shape_id,
#     container="shape",
#     row_index=None,
#     col_index=None
# ):
#     shape = _find_shape_by_id(prs, slide_number, shape_id)

#     if container == "shape":
#         return _get_text_runs_from_shape(shape)

#     if container == "table_cell":
#         if row_index is None or col_index is None:
#             raise ValueError("row_index and col_index required for table_cell.")
#         return _get_text_runs_from_table_cell(shape, row_index, col_index)

#     raise ValueError(f"Unknown container type: {container}")



# def _split_run_by_range(text_range, start, end):
#     """
#     Split a TextRange into [before][target][after] and
#     return the target TextRange.
#     """
#     full_len = len(text_range.Text)

#     if start < 0 or end > full_len or start >= end:
#         raise ValueError("Invalid char range inside run.")

#     # before
#     if start > 0:
#         text_range.Characters(1, start).InsertAfter(
#             text_range.Characters(start + 1, full_len - start).Text
#         )
#         target = text_range.Characters(1, end - start)
#     else:
#         target = text_range.Characters(1, end)

#     # after 제거
#     if end < full_len:
#         target.Characters(end - start + 1, full_len - end).Delete()

#     return target


# def _resolve_insert_position(
#     text: str,
#     preceding_text: str,
#     char_start_index: int,
# ):
#     if preceding_text:
#         idx = text.find(preceding_text)
#         if idx != -1:
#             return idx + len(preceding_text)

#     # fallback: clamp index
#     return max(0, min(len(text), char_start_index))


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

