import win32com.client
import json


# --- Internal Helper Functions ---

def _hex_to_rgb_int(hex_str):
    """Converts HEX string (#FFFFFF or FFFFFF) to win32-compatible BGR integer."""
    hex_str = hex_str.lstrip('#')
    if len(hex_str) != 6:
        raise ValueError("HEX code must be 6 characters long (e.g., FF0000).")
    
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    # Win32com uses BGR structure: (b << 16) | (g << 8) | r
    return (b << 16) | (g << 8) | r

def find_shape_by_id(prs, slide_index, shape_id):
    """Finds a specific Shape object by its unique ID on a given slide."""
    try:
        # Note: slide_index usually starts from 1 in PPT API
        slide = prs.Slides(slide_index)
        for shape in slide.Shapes:
            if shape.Id == shape_id:
                return shape
    except Exception as e:
        raise ValueError(f"Error accessing slide {slide_index}: {e}")
        
    raise ValueError(f"Shape with ID {shape_id} not found on slide {slide_index}.")


# --- [A] Text Style Editing ---

def set_text_style(prs, slide_index, shape_id, font_size=None, color_hex=None, bold=None):
    """Modifies the font size, color, and weight of the text in a shape."""
    shape = find_shape_by_id(prs, slide_index, shape_id)
    if not shape.HasTextFrame:
        return f"Error: Shape {shape_id} cannot contain text."
    
    tr = shape.TextFrame.TextRange
    if font_size: tr.Font.Size = font_size
    if bold is not None: tr.Font.Bold = bold
    if color_hex:
        tr.Font.Color.RGB = _hex_to_rgb_int(color_hex)
    
    return f"Successfully updated style for Shape {shape_id}."


# --- [B] Text Content Editing ---

def update_text(prs, slide_index, shape_id, new_text):
    """Updates the actual text content inside a shape."""
    shape = find_shape_by_id(prs, slide_index, shape_id)
    if not shape.HasTextFrame:
        return f"Error: Shape {shape_id} does not support text editing."
    
    shape.TextFrame.TextRange.Text = new_text
    return f"Successfully updated text for Shape {shape_id}."


# --- [C] Layout / Geometry Editing ---

def adjust_layout(prs, slide_index, shape_id, left=None, top=None, width=None, height=None):
    """Adjusts the position (left, top) and size (width, height) of a shape."""
    shape = find_shape_by_id(prs, slide_index, shape_id)
    if left is not None: shape.Left = left
    if top is not None: shape.Top = top
    if width is not None: shape.Width = width
    if height is not None: shape.Height = height
    return f"Successfully adjusted layout for Shape {shape_id}."


# --- [D] Object Lifecycle ---

def manage_object(prs, slide_index, action, shape_id=None, shape_type=1, left=100, top=100, width=100, height=100):
    """Adds a new shape or deletes an existing one."""
    slide = prs.Slides(slide_index)
    if action == "add":
        new_shape = slide.Shapes.AddShape(shape_type, left, top, width, height)
        return f"Shape created successfully (ID: {new_shape.Id})."
    elif action == "delete" and shape_id:
        find_shape_by_id(prs, slide_index, shape_id).Delete()
        return f"Shape {shape_id} deleted successfully."
    return "Invalid action or missing shape_id."


# --- [E] Visual Style / Theme ---

def apply_visual_style(prs, slide_index, shape_id, bg_color_hex=None, line_color_hex=None, line_weight=None):
    """Sets the background fill color, border color, and border weight of a shape."""
    shape = find_shape_by_id(prs, slide_index, shape_id)
    results = []

    if bg_color_hex:
        shape.Fill.Visible = True 
        shape.Fill.ForeColor.RGB = _hex_to_rgb_int(bg_color_hex)
        results.append(f"background({bg_color_hex})")

    if line_color_hex:
        shape.Line.Visible = True 
        shape.Line.ForeColor.RGB = _hex_to_rgb_int(line_color_hex)
        results.append(f"line color({line_color_hex})")

    if line_weight is not None:
        shape.Line.Visible = True
        shape.Line.Weight = line_weight
        results.append(f"line weight({line_weight}pt)")

    if not results:
        return f"No changes applied to Shape {shape_id}."
    
    return f"Shape {shape_id} visual style updated: " + ", ".join(results)


# --- [F] Consistency / Polishing ---

def align_to_object(prs, slide_index, target_id, base_id, side="right"):
    """Aligns the target shape relative to a base shape."""
    target = find_shape_by_id(prs, slide_index, target_id)
    base = find_shape_by_id(prs, slide_index, base_id)
    
    margin = 10 # Default spacing
    
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

def sum_numbers(a: float, b: float) -> dict:
    """Calculates the sum of two numbers."""
    return {"result": a + b}


# --- Tool Schema ---

TOOLS_SCHEMA = [
    {
        "type": "function",
        "function": {
            "name": "set_text_style",
            "description": "Precisely adjusts text styles including font size, color, and bold properties.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer", "description": "1-based index of the slide."},
                    "shape_id": {"type": "integer", "description": "The unique ID of the shape."},
                    "font_size": {"type": "number", "description": "Font size in points (pt)."},
                    "color_hex": {"type": "string", "description": "Hex color code (e.g., FF0000)."},
                    "bold": {"type": "boolean", "description": "Whether to make text bold."}
                },
                "required": ["slide_index", "shape_id"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "update_text",
            "description": "Updates the text content inside a specific shape.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "shape_id": {"type": "integer"},
                    "new_text": {"type": "string", "description": "The new text string to insert."}
                },
                "required": ["slide_index", "shape_id", "new_text"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "adjust_layout",
            "description": "Adjusts the position and dimensions of a shape.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "shape_id": {"type": "integer"},
                    "left": {"type": "number"},
                    "top": {"type": "number"},
                    "width": {"type": "number"},
                    "height": {"type": "number"}
                },
                "required": ["slide_index", "shape_id"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "manage_object",
            "description": "Creates or deletes shapes. (Heart: 21, Star: 9, Rectangle: 1, Oval: 9)",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "action": {"type": "string", "enum": ["add", "delete"]},
                    "shape_id": {"type": "integer", "description": "Required only for delete action."},
                    "shape_type": {"type": "integer", "default": 1},
                    "left": {"type": "number", "default": 100},
                    "top": {"type": "number", "default": 100},
                    "width": {"type": "number", "default": 100},
                    "height": {"type": "number", "default": 100}
                },
                "required": ["slide_index", "action"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "apply_visual_style",
            "description": "Sets the background color, line color, and line weight of a shape.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "shape_id": {"type": "integer"},
                    "bg_color_hex": {"type": "string", "description": "Background HEX code."},
                    "line_color_hex": {"type": "string", "description": "Line/Border HEX code."},
                    "line_weight": {"type": "number", "description": "Line thickness in pt."}
                },
                "required": ["slide_index", "shape_id"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "align_to_object",
            "description": "Aligns a target shape relative to a base shape.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "target_id": {"type": "integer"},
                    "base_id": {"type": "integer"},
                    "side": {"type": "string", "enum": ["left", "right", "top", "bottom", "center"]}
                },
                "required": ["slide_index", "target_id", "base_id", "side"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "sum_numbers",
            "description": "Calculates the sum of two numbers.",
            "parameters": {
                "type": "object",
                "properties": {
                    "a": {"type": "number"},
                    "b": {"type": "number"}
                },
                "required": ["a", "b"]
            }
        }
    }
]

# --- Function Map ---

FUNCTION_MAP = {
    "set_text_style": set_text_style,
    "update_text": update_text,
    "adjust_layout": adjust_layout,
    "manage_object": manage_object,
    "apply_visual_style": apply_visual_style,
    "align_to_object": align_to_object,
    "sum_numbers": sum_numbers,
}