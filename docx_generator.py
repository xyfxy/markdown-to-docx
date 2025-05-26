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
from docx.oxml.ns import qn
# from docx.enum.style import WD_STYLE_TYPE # For creating custom styles if needed later

# --- Font Name Constants ---
FONT_CHINESE_PRIMARY = "仿宋_GB2312"  # Windows系统中的仿宋字体名称
FONT_CHINESE_FALLBACK = "FangSong"    # 备用仿宋字体名称
FONT_CHINESE_FALLBACK2 = "SimSun"     # 最终回退字体
FONT_ENGLISH_NUMBERS = "Times New Roman"
FONT_HEADING_H1 = "仿宋_GB2312"  # 改为仿宋字体
FONT_HEADING_H1_FALLBACK = "FangSong"  # 仿宋回退字体
FONT_HEADING_H1_FALLBACK2 = "SimSun"  # 最终回退
FONT_HEADING_H2 = "黑体"
FONT_HEADING_H2_FALLBACK = "SimHei"
FONT_HEADING_H3 = "楷体_GB2312"       # Windows系统中的楷体字体名称
FONT_HEADING_H3_FALLBACK = "KaiTi"

# Default PRD values - 3号字体约等于16pt
DEFAULT_BODY_FONT_SIZE_PT = 16.0
DEFAULT_LINE_SPACING_PT = 27.0

# 字号对照表（中文字号到磅值的转换）
FONT_SIZE_MAPPING = {
    "二号": 22.0,  # H1标题
    "三号": 16.0,  # 正文和其他标题
}

# 中文数字转换
CHINESE_NUMBERS = ["", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]

def convert_to_chinese_number(num):
    """将阿拉伯数字转换为中文数字"""
    if num <= 0 or num > 10:
        return str(num)  # 超出范围返回原数字
    return CHINESE_NUMBERS[num]

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

        # 层次编号计数器
        self.heading_counters = {
            2: 0,  # 二级标题：一、二、三...
            3: 0,  # 三级标题：（一）（二）（三）...
            4: 0,  # 四级标题：1. 2. 3. ...
            5: 0   # 五级标题：（1）（2）（3）...
        }

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
        # 设置正文字体为仿宋
        self._set_font_with_fallback(font_style, FONT_CHINESE_PRIMARY, FONT_CHINESE_FALLBACK, FONT_CHINESE_FALLBACK2)
        
        # Apply custom body font size to the 'Normal' style's font if provided.
        # This helps ensure consistency if paragraphs are added without explicit run styling.
        font_style.size = Pt(self.body_font_size)

    def _set_font_with_fallback(self, font_obj, primary_font: str, fallback_font: str, fallback_font2: str = None):
        """设置字体，包含多级回退机制，使用XML方式确保WPS兼容性"""
        fonts_to_try = [primary_font, fallback_font]
        if fallback_font2:
            fonts_to_try.append(fallback_font2)
        fonts_to_try.append("SimSun")  # 最终回退
        
        selected_font = None
        for font_name in fonts_to_try:
            try:
                # 先尝试设置常规属性
                font_obj.name = font_name
                selected_font = font_name
                break
            except Exception:
                continue
        
        # 使用XML方式设置东亚字体，确保WPS识别
        if selected_font and hasattr(font_obj, 'element') and hasattr(font_obj.element, 'rPr'):
            try:
                from docx.oxml import OxmlElement
                rPr = font_obj.element.rPr
                
                # 获取或创建rFonts元素
                rFonts = rPr.rFonts
                if rFonts is None:
                    rFonts = OxmlElement('w:rFonts')
                    rPr.append(rFonts)
                
                # 设置字体
                rFonts.set(qn('w:eastAsia'), selected_font)
                rFonts.set(qn('w:ascii'), FONT_ENGLISH_NUMBERS)
                rFonts.set(qn('w:hAnsi'), FONT_ENGLISH_NUMBERS)
                
            except Exception as e:
                # 回退到标准方式
                try:
                    font_obj.name_east_asia = selected_font
                    font_obj.name_ascii = FONT_ENGLISH_NUMBERS
                except Exception:
                    pass

    def _apply_run_font_style(self, run, font_type: str, base_font_chinese: str, base_font_chinese_fallback: str, size_pt: float, bold: bool = False, italic: bool = False):
        run.font.size = Pt(size_pt)
        if font_type == 'chinese':
            # 为中文字体提供更好的回退机制
            if base_font_chinese == FONT_CHINESE_PRIMARY:
                self._set_font_with_fallback(run.font, base_font_chinese, base_font_chinese_fallback, FONT_CHINESE_FALLBACK2)
            elif base_font_chinese == FONT_HEADING_H1:
                self._set_font_with_fallback(run.font, base_font_chinese, FONT_HEADING_H1_FALLBACK, FONT_HEADING_H1_FALLBACK2)
            else:
                self._set_font_with_fallback(run.font, base_font_chinese, base_font_chinese_fallback)
        else: # 'english_number'
            # 英文数字使用Times New Roman，但东亚字体保持与中文一致
            run.font.name = FONT_ENGLISH_NUMBERS
            
            # 使用XML方式设置字体，确保WPS兼容性
            if hasattr(run.font, 'element') and hasattr(run.font.element, 'rPr'):
                try:
                    from docx.oxml import OxmlElement
                    rPr = run.font.element.rPr
                    
                    # 获取或创建rFonts元素
                    rFonts = rPr.rFonts
                    if rFonts is None:
                        rFonts = OxmlElement('w:rFonts')
                        rPr.append(rFonts)
                    
                    # 设置字体
                    rFonts.set(qn('w:ascii'), FONT_ENGLISH_NUMBERS)
                    rFonts.set(qn('w:hAnsi'), FONT_ENGLISH_NUMBERS)
                    # 东亚字体使用中文字体，确保混排一致性
                    if base_font_chinese == FONT_CHINESE_PRIMARY:
                        rFonts.set(qn('w:eastAsia'), FONT_CHINESE_PRIMARY)
                    else:
                        rFonts.set(qn('w:eastAsia'), base_font_chinese)
                        
                except Exception:
                    # 回退到标准方式
                    try:
                        run.font.name_ascii = FONT_ENGLISH_NUMBERS
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
        """添加主标题（H1）- 二号方正小标宋简体，居中，不加粗，行距32磅"""
        p = self.document.add_paragraph()
        p_format = p.paragraph_format
        p_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        p_format.line_spacing = Pt(32) # Fixed line spacing for H1
        segments = segment_text_for_font_styling(text)
        for segment_text, font_type in segments:
            run = p.add_run(segment_text)
            # 使用二号字体（22pt），不加粗
            self._apply_run_font_style(run, font_type, FONT_HEADING_H1, FONT_HEADING_H1_FALLBACK, FONT_SIZE_MAPPING["二号"], bold=False)
        # 标题下空两行
        self.document.add_paragraph()
        self.document.add_paragraph()

    def _add_h2(self, text_content: str):
        """添加二级标题 - 黑体，使用"一、" """
        self.heading_counters[2] += 1
        # 重置下级标题计数器
        self.heading_counters[3] = 0
        self.heading_counters[4] = 0
        self.heading_counters[5] = 0
        
        chinese_num = convert_to_chinese_number(self.heading_counters[2])
        full_text = f"{chinese_num}、{text_content}"
        p = self.document.add_paragraph()
        segments = segment_text_for_font_styling(full_text)
        for segment_text, font_type in segments:
            run = p.add_run(segment_text)
            # 使用三号黑体
            self._apply_run_font_style(run, font_type, FONT_HEADING_H2, FONT_HEADING_H2_FALLBACK, FONT_SIZE_MAPPING["三号"])

    def _add_h3(self, text_content: str):
        """添加三级标题 - 楷体，使用"（一）" """
        self.heading_counters[3] += 1
        # 重置下级标题计数器
        self.heading_counters[4] = 0
        self.heading_counters[5] = 0
        
        chinese_num = convert_to_chinese_number(self.heading_counters[3])
        full_text = f"（{chinese_num}）{text_content}"
        p = self.document.add_paragraph()
        segments = segment_text_for_font_styling(full_text)
        for segment_text, font_type in segments:
            run = p.add_run(segment_text)
            # 使用三号楷体
            self._apply_run_font_style(run, font_type, FONT_HEADING_H3, FONT_HEADING_H3_FALLBACK, FONT_SIZE_MAPPING["三号"])

    def _add_h4(self, text_content: str):
        """添加四级标题 - 仿宋体，使用"1." """
        self.heading_counters[4] += 1
        # 重置下级标题计数器
        self.heading_counters[5] = 0
        
        full_text = f"{self.heading_counters[4]}.{text_content}"
        p = self.document.add_paragraph()
        segments = segment_text_for_font_styling(full_text)
        for segment_text, font_type in segments:
            run = p.add_run(segment_text)
            # 使用三号仿宋体
            self._apply_run_font_style(run, font_type, FONT_CHINESE_PRIMARY, FONT_CHINESE_FALLBACK, FONT_SIZE_MAPPING["三号"])
            
    def _add_h5(self, text_content: str):
        """添加五级标题 - 仿宋体，使用"（1）" """
        self.heading_counters[5] += 1
        
        full_text = f"（{self.heading_counters[5]}）{text_content}"
        p = self.document.add_paragraph()
        segments = segment_text_for_font_styling(full_text)
        for segment_text, font_type in segments:
            run = p.add_run(segment_text)
            # 使用三号仿宋体
            self._apply_run_font_style(run, font_type, FONT_CHINESE_PRIMARY, FONT_CHINESE_FALLBACK, FONT_SIZE_MAPPING["三号"])

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
