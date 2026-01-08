import win32com.client
import json
from typing import Optional, Union, List, Dict, Any


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


def find_shape_by_id(prs, slide_index, shape_id):
    """Finds a specific Shape object by its unique ID on a given slide."""
    try:
        slide = prs.Slides(slide_index)
        for shape in slide.Shapes:
            if shape.Id == shape_id:
                return shape
    except Exception as e:
        raise ValueError(f"Error accessing slide {slide_index}: {e}")
    raise ValueError(f"Shape with ID {shape_id} not found on slide {slide_index}.")


def find_shapes_by_name(prs, slide_index, shape_name):
    """Finds shapes by name pattern on a given slide."""
    slide = prs.Slides(slide_index)
    shapes = []
    for shape in slide.Shapes:
        if shape_name.lower() in shape.Name.lower():
            shapes.append(shape)
    return shapes


# --- [A] Enhanced Text Style Editing ---

def set_text_style(prs, slide_index, shape_id, font_size=None, color_hex=None, 
                   bold=None, italic=None, underline=None, font_name=None,
                   paragraph_index=None, char_start=None, char_end=None):
    """
    Modifies text styles with support for partial text selection.
    
    Args:
        paragraph_index: Specific paragraph to style (0-based, None = all)
        char_start, char_end: Character range within text (None = all)
    """
    shape = find_shape_by_id(prs, slide_index, shape_id)
    if not shape.HasTextFrame:
        return f"Error: Shape {shape_id} cannot contain text."
    
    # Select target text range
    if paragraph_index is not None:
        tr = shape.TextFrame.TextRange.Paragraphs(paragraph_index + 1)
    elif char_start is not None and char_end is not None:
        tr = shape.TextFrame.TextRange.Characters(char_start, char_end - char_start)
    else:
        tr = shape.TextFrame.TextRange
    
    # Apply styles
    if font_size: tr.Font.Size = font_size
    if bold is not None: tr.Font.Bold = bold
    if italic is not None: tr.Font.Italic = italic
    if underline is not None: tr.Font.Underline = underline
    if font_name: tr.Font.Name = font_name
    if color_hex: tr.Font.Color.RGB = _hex_to_rgb_int(color_hex)
    
    return f"Successfully updated style for Shape {shape_id}."


def set_paragraph_alignment(prs, slide_index, shape_id, alignment="left", 
                           line_spacing=None, space_before=None, space_after=None):
    """
    Adjusts paragraph-level formatting.
    
    Args:
        alignment: 'left', 'center', 'right', 'justify', 'distribute'
        line_spacing: Line spacing multiplier (e.g., 1.5)
        space_before/after: Points before/after paragraph
    """
    shape = find_shape_by_id(prs, slide_index, shape_id)
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


def add_bullet_points(prs, slide_index, shape_id, bullet_type="bullet", 
                     bullet_char=None, start_value=1):
    """
    Adds or modifies bullet/numbering format.
    
    Args:
        bullet_type: 'bullet', 'number', 'none'
        bullet_char: Custom bullet character
        start_value: Starting number for numbered lists
    """
    shape = find_shape_by_id(prs, slide_index, shape_id)
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

def update_text(prs, slide_index, shape_id, new_text, append=False, 
                paragraph_index=None):
    """
    Updates text content with append option and paragraph targeting.
    
    Args:
        append: If True, adds to existing text instead of replacing
        paragraph_index: Target specific paragraph (0-based)
    """
    shape = find_shape_by_id(prs, slide_index, shape_id)
    if not shape.HasTextFrame:
        return f"Error: Shape {shape_id} does not support text editing."
    
    if paragraph_index is not None:
        para = shape.TextFrame.TextRange.Paragraphs(paragraph_index + 1)
        if append:
            para.Text = para.Text + new_text
        else:
            para.Text = new_text
    else:
        if append:
            shape.TextFrame.TextRange.Text += new_text
        else:
            shape.TextFrame.TextRange.Text = new_text
    
    return f"Successfully updated text for Shape {shape_id}."


def find_and_replace(prs, slide_index, shape_id, find_text, replace_text, 
                    match_case=False):
    """Finds and replaces text within a shape."""
    shape = find_shape_by_id(prs, slide_index, shape_id)
    if not shape.HasTextFrame:
        return f"Error: Shape {shape_id} does not support text editing."
    
    current_text = shape.TextFrame.TextRange.Text
    
    if match_case:
        new_text = current_text.replace(find_text, replace_text)
    else:
        # Case-insensitive replace
        import re
        new_text = re.sub(re.escape(find_text), replace_text, current_text, flags=re.IGNORECASE)
    
    shape.TextFrame.TextRange.Text = new_text
    count = len(re.findall(re.escape(find_text), current_text, re.IGNORECASE if not match_case else 0))
    
    return f"Replaced {count} occurrence(s) in Shape {shape_id}."


# --- [C] Enhanced Layout / Geometry Editing ---

def adjust_layout(prs, slide_index, shape_id, left=None, top=None, 
                 width=None, height=None, rotation=None):
    """Adjusts position, size, and rotation of a shape."""
    shape = find_shape_by_id(prs, slide_index, shape_id)
    
    if left is not None: shape.Left = left
    if top is not None: shape.Top = top
    if width is not None: shape.Width = width
    if height is not None: shape.Height = height
    if rotation is not None: shape.Rotation = rotation
    
    return f"Successfully adjusted layout for Shape {shape_id}."


def distribute_shapes(prs, slide_index, shape_ids, direction="horizontal", 
                     spacing=None):
    """
    Distributes multiple shapes evenly.
    
    Args:
        direction: 'horizontal' or 'vertical'
        spacing: Fixed spacing between shapes (if None, distribute evenly)
    """
    shapes = [find_shape_by_id(prs, slide_index, sid) for sid in shape_ids]
    
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


def align_shapes(prs, slide_index, shape_ids, align_type="left"):
    """
    Aligns multiple shapes to each other.
    
    Args:
        align_type: 'left', 'right', 'top', 'bottom', 'center_h', 'center_v'
    """
    shapes = [find_shape_by_id(prs, slide_index, sid) for sid in shape_ids]
    
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

def manage_object(prs, slide_index, action, shape_id=None, shape_type=1, 
                 left=100, top=100, width=100, height=100, text=None):
    """Creates, deletes, or duplicates shapes."""
    slide = prs.Slides(slide_index)
    
    if action == "add":
        new_shape = slide.Shapes.AddShape(shape_type, left, top, width, height)
        if text and new_shape.HasTextFrame:
            new_shape.TextFrame.TextRange.Text = text
        return f"Shape created successfully (ID: {new_shape.Id})."
    
    elif action == "delete" and shape_id:
        find_shape_by_id(prs, slide_index, shape_id).Delete()
        return f"Shape {shape_id} deleted successfully."
    
    elif action == "duplicate" and shape_id:
        original = find_shape_by_id(prs, slide_index, shape_id)
        duplicate = original.Duplicate()
        duplicate.Left += 20  # Offset slightly
        duplicate.Top += 20
        return f"Shape {shape_id} duplicated (New ID: {duplicate.Id})."
    
    return "Invalid action or missing shape_id."


def add_textbox(prs, slide_index, left, top, width, height, text=""):
    """Creates a new textbox with specified text."""
    slide = prs.Slides(slide_index)
    textbox = slide.Shapes.AddTextbox(1, left, top, width, height)  # msoTextOrientationHorizontal
    if text:
        textbox.TextFrame.TextRange.Text = text
    return f"Textbox created (ID: {textbox.Id})."


def add_image(prs, slide_index, image_path, left, top, width=None, height=None):
    """Inserts an image onto the slide."""
    slide = prs.Slides(slide_index)
    
    if width and height:
        picture = slide.Shapes.AddPicture(image_path, False, True, left, top, width, height)
    else:
        picture = slide.Shapes.AddPicture(image_path, False, True, left, top)
    
    return f"Image inserted (ID: {picture.Id})."


def group_shapes(prs, slide_index, shape_ids):
    """Groups multiple shapes together."""
    slide = prs.Slides(slide_index)
    shapes = [find_shape_by_id(prs, slide_index, sid) for sid in shape_ids]
    
    # Create shape range
    shape_range = slide.Shapes.Range([s.Id for s in shapes])
    grouped = shape_range.Group()
    
    return f"Grouped {len(shape_ids)} shapes (Group ID: {grouped.Id})."


def ungroup_shapes(prs, slide_index, group_id):
    """Ungroups a grouped shape."""
    group = find_shape_by_id(prs, slide_index, group_id)
    ungrouped = group.Ungroup()
    
    return f"Ungrouped shape {group_id} into {ungrouped.Count} shapes."


# --- [E] Enhanced Visual Style / Theme ---

def apply_visual_style(prs, slide_index, shape_id, bg_color_hex=None, 
                      line_color_hex=None, line_weight=None, line_style=None,
                      transparency=None, shadow=None):
    """
    Sets comprehensive visual styles.
    
    Args:
        line_style: 'solid', 'dash', 'dot', 'dash_dot'
        transparency: 0-1 (0=opaque, 1=transparent)
        shadow: True/False to enable/disable shadow
    """
    shape = find_shape_by_id(prs, slide_index, shape_id)
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


def apply_gradient_fill(prs, slide_index, shape_id, color1_hex, color2_hex, 
                       gradient_type="linear", angle=0):
    """
    Applies gradient fill to a shape.
    
    Args:
        gradient_type: 'linear', 'radial', 'rectangular', 'path'
        angle: Gradient angle in degrees (for linear)
    """
    shape = find_shape_by_id(prs, slide_index, shape_id)
    
    gradient_map = {'linear': 1, 'radial': 3, 'rectangular': 4, 'path': 5}
    
    shape.Fill.TwoColorGradient(gradient_map.get(gradient_type, 1), 1)
    shape.Fill.ForeColor.RGB = _hex_to_rgb_int(color1_hex)
    shape.Fill.BackColor.RGB = _hex_to_rgb_int(color2_hex)
    
    if gradient_type == "linear":
        # Set gradient angle
        shape.Fill.GradientAngle = angle
    
    return f"Gradient applied to Shape {shape_id}."


def set_shape_effect(prs, slide_index, shape_id, effect_type, **kwargs):
    """
    Applies special effects to shapes.
    
    Args:
        effect_type: 'glow', 'soft_edge', 'reflection', '3d'
        kwargs: Effect-specific parameters
    """
    shape = find_shape_by_id(prs, slide_index, shape_id)
    
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

def align_to_object(prs, slide_index, target_id, base_id, side="right", margin=10):
    """Aligns the target shape relative to a base shape with custom margin."""
    target = find_shape_by_id(prs, slide_index, target_id)
    base = find_shape_by_id(prs, slide_index, base_id)
    
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


def match_formatting(prs, slide_index, source_id, target_ids):
    """Copies formatting from source shape to target shapes."""
    source = find_shape_by_id(prs, slide_index, source_id)
    targets = [find_shape_by_id(prs, slide_index, tid) for tid in target_ids]
    
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


def set_z_order(prs, slide_index, shape_id, order="bring_to_front"):
    """
    Changes the z-order (layering) of a shape.
    
    Args:
        order: 'bring_to_front', 'send_to_back', 'bring_forward', 'send_backward'
    """
    shape = find_shape_by_id(prs, slide_index, shape_id)
    
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


def delete_slide(prs, slide_index):
    """Deletes a specific slide."""
    prs.Slides(slide_index).Delete()
    return f"Slide {slide_index} deleted."


def duplicate_slide(prs, slide_index):
    """Duplicates a specific slide."""
    original = prs.Slides(slide_index)
    duplicate = original.Duplicate()
    return f"Slide {slide_index} duplicated to position {duplicate.SlideIndex}."


# --- [H] Table Operations ---

def add_table(prs, slide_index, rows, cols, left, top, width, height):
    """Creates a table on the slide."""
    slide = prs.Slides(slide_index)
    table = slide.Shapes.AddTable(rows, cols, left, top, width, height)
    return f"Table created (ID: {table.Id}, {rows}x{cols})."


def update_table_cell(prs, slide_index, table_id, row, col, text, 
                     font_size=None, color_hex=None, bg_color_hex=None):
    """Updates content and style of a specific table cell."""
    table_shape = find_shape_by_id(prs, slide_index, table_id)
    
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

def add_animation(prs, slide_index, shape_id, effect_type="appear", 
                 trigger="on_click", duration=1.0):
    """
    Adds animation to a shape.
    
    Args:
        effect_type: 'appear', 'fade', 'fly_in', 'zoom', etc.
        trigger: 'on_click', 'with_previous', 'after_previous'
        duration: Animation duration in seconds
    """
    slide = prs.Slides(slide_index)
    shape = find_shape_by_id(prs, slide_index, shape_id)
    
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


def set_slide_transition(prs, slide_index, transition_type="fade", 
                        duration=1.0, advance_on_time=None):
    """
    Sets slide transition effect.
    
    Args:
        transition_type: 'fade', 'push', 'wipe', 'split', etc.
        duration: Transition duration in seconds
        advance_on_time: Auto-advance after N seconds (None = manual)
    """
    slide = prs.Slides(slide_index)
    
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
    
    return f"Transition '{transition_type}' applied to slide {slide_index}."


# --- Complete Tool Schema ---
TOOLS_SCHEMA = [

    # =========================
    # [A] Text Styling
    # =========================
    {
        "type": "function",
        "function": {
            "name": "set_text_style",
            "description": "Precisely adjusts text styles including font, size, color, bold, italic, underline. Supports partial text selection.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "shape_id": {"type": "integer"},
                    "font_size": {"type": "number"},
                    "color_hex": {"type": "string"},
                    "bold": {"type": "boolean"},
                    "italic": {"type": "boolean"},
                    "underline": {"type": "boolean"},
                    "font_name": {"type": "string"},
                    "paragraph_index": {"type": "integer"},
                    "char_start": {"type": "integer"},
                    "char_end": {"type": "integer"}
                },
                "required": ["slide_index", "shape_id"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "set_paragraph_alignment",
            "description": "Sets paragraph alignment and spacing options.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "shape_id": {"type": "integer"},
                    "alignment": {
                        "type": "string",
                        "enum": ["left", "center", "right", "justify", "distribute"]
                    },
                    "line_spacing": {"type": "number"},
                    "space_before": {"type": "number"},
                    "space_after": {"type": "number"}
                },
                "required": ["slide_index", "shape_id"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "add_bullet_points",
            "description": "Applies bullet or numbering styles to text.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "shape_id": {"type": "integer"},
                    "bullet_type": {
                        "type": "string",
                        "enum": ["bullet", "number", "none"]
                    },
                    "bullet_char": {"type": "string"},
                    "start_value": {"type": "integer"}
                },
                "required": ["slide_index", "shape_id"]
            }
        }
    },

    # =========================
    # [B] Text Content
    # =========================
    {
        "type": "function",
        "function": {
            "name": "update_text",
            "description": "Updates or appends text in a shape, optionally targeting a paragraph.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "shape_id": {"type": "integer"},
                    "new_text": {"type": "string"},
                    "append": {"type": "boolean"},
                    "paragraph_index": {"type": "integer"}
                },
                "required": ["slide_index", "shape_id", "new_text"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "find_and_replace",
            "description": "Finds and replaces text within a shape.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "shape_id": {"type": "integer"},
                    "find_text": {"type": "string"},
                    "replace_text": {"type": "string"},
                    "match_case": {"type": "boolean"}
                },
                "required": ["slide_index", "shape_id", "find_text", "replace_text"]
            }
        }
    },

    # =========================
    # [C] Layout / Geometry
    # =========================
    {
        "type": "function",
        "function": {
            "name": "adjust_layout",
            "description": "Adjusts shape position, size, and rotation.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "shape_id": {"type": "integer"},
                    "left": {"type": "number"},
                    "top": {"type": "number"},
                    "width": {"type": "number"},
                    "height": {"type": "number"},
                    "rotation": {"type": "number"}
                },
                "required": ["slide_index", "shape_id"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "distribute_shapes",
            "description": "Distributes multiple shapes evenly.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "shape_ids": {
                        "type": "array",
                        "items": {"type": "integer"}
                    },
                    "direction": {
                        "type": "string",
                        "enum": ["horizontal", "vertical"]
                    },
                    "spacing": {"type": "number"}
                },
                "required": ["slide_index", "shape_ids"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "align_shapes",
            "description": "Aligns multiple shapes relative to each other.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "shape_ids": {
                        "type": "array",
                        "items": {"type": "integer"}
                    },
                    "align_type": {
                        "type": "string",
                        "enum": ["left", "right", "top", "bottom", "center_h", "center_v"]
                    }
                },
                "required": ["slide_index", "shape_ids"]
            }
        }
    },

    # =========================
    # [D] Object Lifecycle
    # =========================
    {
        "type": "function",
        "function": {
            "name": "manage_object",
            "description": "Creates, deletes, or duplicates a shape.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "action": {
                        "type": "string",
                        "enum": ["add", "delete", "duplicate"]
                    },
                    "shape_id": {"type": "integer"},
                    "shape_type": {"type": "integer"},
                    "left": {"type": "number"},
                    "top": {"type": "number"},
                    "width": {"type": "number"},
                    "height": {"type": "number"},
                    "text": {"type": "string"}
                },
                "required": ["slide_index", "action"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "add_textbox",
            "description": "Adds a textbox to a slide.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "left": {"type": "number"},
                    "top": {"type": "number"},
                    "width": {"type": "number"},
                    "height": {"type": "number"},
                    "text": {"type": "string"}
                },
                "required": ["slide_index", "left", "top", "width", "height"]
            }
        }
    },

    # =========================
    # [E] Visual Style
    # =========================
    {
        "type": "function",
        "function": {
            "name": "apply_visual_style",
            "description": "Applies fill, line, transparency, and shadow styles.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "shape_id": {"type": "integer"},
                    "bg_color_hex": {"type": "string"},
                    "line_color_hex": {"type": "string"},
                    "line_weight": {"type": "number"},
                    "line_style": {
                        "type": "string",
                        "enum": ["solid", "dash", "dot", "dash_dot"]
                    },
                    "transparency": {"type": "number"},
                    "shadow": {"type": "boolean"}
                },
                "required": ["slide_index", "shape_id"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "apply_gradient_fill",
            "description": "Applies a gradient fill to a shape.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "shape_id": {"type": "integer"},
                    "color1_hex": {"type": "string"},
                    "color2_hex": {"type": "string"},
                    "gradient_type": {
                        "type": "string",
                        "enum": ["linear", "radial", "rectangular", "path"]
                    },
                    "angle": {"type": "number"}
                },
                "required": ["slide_index", "shape_id", "color1_hex", "color2_hex"]
            }
        }
    },

    # =========================
    # [F] Slide Management
    # =========================
    {
        "type": "function",
        "function": {
            "name": "add_slide",
            "description": "Adds a new slide.",
            "parameters": {
                "type": "object",
                "properties": {
                    "layout_index": {"type": "integer"},
                    "position": {"type": "integer"}
                }
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "delete_slide",
            "description": "Deletes a slide.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"}
                },
                "required": ["slide_index"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "duplicate_slide",
            "description": "Duplicates a slide.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"}
                },
                "required": ["slide_index"]
            }
        }
    }
]


FUNCTION_MAP = {
    "set_text_style": set_text_style,
    "set_paragraph_alignment": set_paragraph_alignment,
    "add_bullet_points": add_bullet_points,
    "update_text": update_text,
    "find_and_replace": find_and_replace,
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


