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

# def update_text(prs, slide_index, shape_id, new_text):
#     """Updates the actual text content inside a shape."""
#     shape = find_shape_by_id(prs, slide_index, shape_id)
#     if not shape.HasTextFrame:
#         return f"Error: Shape {shape_id} does not support text editing."
    
#     shape.TextFrame.TextRange.Text = new_text
#     return f"Successfully updated text for Shape {shape_id}."


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
