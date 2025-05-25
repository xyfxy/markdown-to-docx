# docx_generator.py
"""
This module provides the DocxWriter class for creating and populating .docx files
based on structured content, with specific formatting for Chinese documents.
"""
import re
import typing # Added for Optional type hint

from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
# from docx.enum.style import WD_STYLE_TYPE # For creating custom styles if needed later

# --- Font Name Constants ---
FONT_CHINESE_PRIMARY = "仿宋"
FONT_CHINESE_FALLBACK = "SimSun"
FONT_ENGLISH_NUMBERS = "Times New Roman"
FONT_HEADING_H1 = "方正小标宋简体"
FONT_HEADING_H1_FALLBACK = "SimHei"
FONT_HEADING_H2 = "黑体"
FONT_HEADING_H2_FALLBACK = "SimHei"
FONT_HEADING_H3 = "楷体"
FONT_HEADING_H3_FALLBACK = "KaiTi"

# Default PRD values
DEFAULT_BODY_FONT_SIZE_PT = 16.0
DEFAULT_LINE_SPACING_PT = 27.0


def is_chinese_char(char):
    return '\u4e00' <= char <= '\u9fff'

def segment_text_for_font_styling(text: str) -> list[tuple[str, str]]:
    segments = []
    if not text:
        return segments
    current_segment = ""
    current_type = 'chinese' if is_chinese_char(text[0]) else 'english_number'
    for char in text:
        char_type = 'chinese' if is_chinese_char(char) else 'english_number'
        if char_type == current_type:
            current_segment += char
        else:
            segments.append((current_segment, current_type))
            current_segment = char
            current_type = char_type
    if current_segment:
        segments.append((current_segment, current_type))
    return segments


class DocxWriter:
    def __init__(self, 
                 output_docx_path: str = "output.docx",
                 line_spacing_pt: typing.Optional[float] = None,
                 font_size_pt: typing.Optional[float] = None,
                 template_path: typing.Optional[str] = None): # Added template_path
        
        self.output_docx_path = output_docx_path
        # If a template is provided, python-docx uses it as the basis for the new document.
        # Styles and some properties can be inherited from the template.
        if template_path:
            try:
                self.document = Document(template_path)
                print(f"Info: Using document template from '{template_path}'")
            except Exception as e:
                print(f"Warning: Could not load template '{template_path}'. Using default. Error: {e}")
                self.document = Document()
        else:
            self.document = Document()

        # Store custom formatting options
        self.custom_line_spacing_pt = line_spacing_pt
        self.custom_font_size_pt = font_size_pt
        
        # Body font size to use (either custom or default)
        self.body_font_size = self.custom_font_size_pt if self.custom_font_size_pt is not None else DEFAULT_BODY_FONT_SIZE_PT

        self._setup_page() # Page setup should apply even with a template, might override template settings
        self._apply_default_styles()


    def _setup_page(self):
        section = self.document.sections[0]
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(3.7)
        section.bottom_margin = Cm(3.5)
        section.left_margin = Cm(2.8)
        section.right_margin = Cm(2.6)

    def _apply_default_styles(self):
        style = self.document.styles['Normal']
        p_format = style.paragraph_format
        p_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        # Apply custom line spacing if provided, else default
        actual_line_spacing = self.custom_line_spacing_pt if self.custom_line_spacing_pt is not None else DEFAULT_LINE_SPACING_PT
        p_format.line_spacing = Pt(actual_line_spacing)
        
        font_style = style.font
        try:
            font_style.name = FONT_CHINESE_PRIMARY
            font_style.name_east_asia = FONT_CHINESE_PRIMARY
        except Exception: # Should be rare for name assignment
            font_style.name = FONT_CHINESE_FALLBACK
            font_style.name_east_asia = FONT_CHINESE_FALLBACK
        font_style.name_ascii = FONT_ENGLISH_NUMBERS
        
        # Apply custom body font size to the 'Normal' style's font if provided.
        # This helps ensure consistency if paragraphs are added without explicit run styling.
        font_style.size = Pt(self.body_font_size)


    def _apply_run_font_style(self, run, font_type: str, base_font_chinese: str, base_font_chinese_fallback: str, size_pt: float, bold: bool = False, italic: bool = False):
        run.font.size = Pt(size_pt)
        if font_type == 'chinese':
            try:
                run.font.name = base_font_chinese
                run.font.name_east_asia = base_font_chinese
            except Exception:
                run.font.name = base_font_chinese_fallback
                run.font.name_east_asia = base_font_chinese_fallback
            run.font.name_ascii = FONT_ENGLISH_NUMBERS
        else: # 'english_number'
            run.font.name = FONT_ENGLISH_NUMBERS
            run.font.name_ascii = FONT_ENGLISH_NUMBERS
            try:
                run.font.name_east_asia = base_font_chinese
            except Exception:
                run.font.name_east_asia = base_font_chinese_fallback
        if bold:
            run.font.bold = True
        if italic:
            run.font.italic = True

    def add_recipient(self, text: str):
        full_recipient_text = text + "："
        p = self.document.add_paragraph()
        p_format = p.paragraph_format
        p_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_format.first_line_indent = Pt(0)
        segments = segment_text_for_font_styling(full_recipient_text)
        for segment_text, font_type in segments:
            run = p.add_run(segment_text)
            self._apply_run_font_style(run, font_type, FONT_CHINESE_PRIMARY, FONT_CHINESE_FALLBACK, self.body_font_size)
        return p

    def add_paragraph(self, text: str):
        p = self.document.add_paragraph()
        p_format = p.paragraph_format
        p_format.first_line_indent = Pt(32) # Standard 2-char indent for 16pt font
        # Adjust indent if font size changes significantly?
        # For now, 32pt is based on 16pt. If body_font_size is e.g. 12pt, indent might be 24pt.
        # A common way is 2 * body_font_size.
        p_format.first_line_indent = Pt(2 * self.body_font_size)


        segments = segment_text_for_font_styling(text)
        for segment_text, font_type in segments:
            run = p.add_run(segment_text)
            self._apply_run_font_style(run, font_type, FONT_CHINESE_PRIMARY, FONT_CHINESE_FALLBACK, self.body_font_size)
        return p

    def add_emphasis(self, text: str, style: str):
        p = self.document.add_paragraph()
        p_format = p.paragraph_format
        p_format.first_line_indent = Pt(2 * self.body_font_size) # Consistent indent
        is_bold = (style == 'bold')
        is_italic = (style == 'italic')
        segments = segment_text_for_font_styling(text)
        for segment_text, font_type in segments:
            run = p.add_run(segment_text)
            self._apply_run_font_style(run, font_type, FONT_CHINESE_PRIMARY, FONT_CHINESE_FALLBACK, self.body_font_size, bold=is_bold, italic=is_italic)
        return p

    # --- Heading Methods ---
    # Headings retain their PRD-specified fixed sizes, not affected by CLI --font-size for body text.
    def _add_main_heading(self, text: str):
        p = self.document.add_paragraph()
        p_format = p.paragraph_format
        p_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        p_format.line_spacing = Pt(32) # Fixed line spacing for H1
        segments = segment_text_for_font_styling(text)
        for segment_text, font_type in segments:
            run = p.add_run(segment_text)
            self._apply_run_font_style(run, font_type, FONT_HEADING_H1, FONT_HEADING_H1_FALLBACK, 22) # Fixed 22pt
        self.document.add_paragraph()
        self.document.add_paragraph()

    def _add_h2(self, text_content: str):
        full_text = "一、 " + text_content
        p = self.document.add_paragraph()
        segments = segment_text_for_font_styling(full_text)
        for segment_text, font_type in segments:
            run = p.add_run(segment_text)
            self._apply_run_font_style(run, font_type, FONT_HEADING_H2, FONT_HEADING_H2_FALLBACK, 16) # Fixed 16pt

    def _add_h3(self, text_content: str):
        full_text = "（一） " + text_content
        p = self.document.add_paragraph()
        segments = segment_text_for_font_styling(full_text)
        for segment_text, font_type in segments:
            run = p.add_run(segment_text)
            self._apply_run_font_style(run, font_type, FONT_HEADING_H3, FONT_HEADING_H3_FALLBACK, 16) # Fixed 16pt

    def _add_h4(self, text_content: str):
        full_text = "1. " + text_content
        p = self.document.add_paragraph()
        segments = segment_text_for_font_styling(full_text)
        for segment_text, font_type in segments:
            run = p.add_run(segment_text)
            self._apply_run_font_style(run, font_type, FONT_CHINESE_PRIMARY, FONT_CHINESE_FALLBACK, 16) # Fixed 16pt
            
    def _add_h5(self, text_content: str):
        full_text = "（1） " + text_content
        p = self.document.add_paragraph()
        segments = segment_text_for_font_styling(full_text)
        for segment_text, font_type in segments:
            run = p.add_run(segment_text)
            self._apply_run_font_style(run, font_type, FONT_CHINESE_PRIMARY, FONT_CHINESE_FALLBACK, 16) # Fixed 16pt

    def add_heading(self, text: str, level: int):
        if level == 1: self._add_main_heading(text)
        elif level == 2: self._add_h2(text)
        elif level == 3: self._add_h3(text)
        elif level == 4: self._add_h4(text)
        elif level == 5: self._add_h5(text)
        else: self.add_paragraph(text) # Fallback with body font size

    def add_list_item(self, text: str, ordered: bool, level: int = 0, item_prefix: str = ""):
        p = self.document.add_paragraph()
        p_format = p.paragraph_format
        indent_unit_cm = 0.75 
        p_format.left_indent = Cm(indent_unit_cm * (level + 1))
        p_format.first_line_indent = Cm(-indent_unit_cm)
        
        if item_prefix:
            prefix_run = p.add_run(item_prefix)
            self._apply_run_font_style(prefix_run, 'english_number', FONT_CHINESE_PRIMARY, FONT_CHINESE_FALLBACK, self.body_font_size)

        segments = segment_text_for_font_styling(text)
        for segment_text, font_type in segments:
            run = p.add_run(segment_text)
            self._apply_run_font_style(run, font_type, FONT_CHINESE_PRIMARY, FONT_CHINESE_FALLBACK, self.body_font_size)
        return p

    def save(self):
        try:
            self.document.save(self.output_docx_path)
        except Exception as e:
            print(f"Error saving document to {self.output_docx_path}: {e}")
            raise
```
