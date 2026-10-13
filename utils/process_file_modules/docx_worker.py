from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX, WD_LINE_SPACING
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from copy import deepcopy
from io import BytesIO
import re
import os
import shutil
import zipfile
from typing import List, Dict, Any, Optional, Tuple, Union, Callable
from dataclasses import dataclass, field

from Site.utils.text.random import text_random
from mysite.logger import logger


@dataclass
class RunStyle:
    text: str = ""
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[Union[bool, int]] = None
    font_name: Optional[str] = None
    font_size: Optional[float] = None
    font_color: Optional[Tuple[int, int, int]] = None
    highlight_color: Optional[int] = None
    all_caps: Optional[bool] = None
    small_caps: Optional[bool] = None
    strike: Optional[bool] = None
    double_strike: Optional[bool] = None
    subscript: Optional[bool] = None
    superscript: Optional[bool] = None
    shadow: Optional[bool] = None
    outline: Optional[bool] = None
    emboss: Optional[bool] = None
    imprint: Optional[bool] = None
    character_style: Optional[str] = None


@dataclass
class ParagraphStyle:
    style_name: Optional[str] = None
    alignment: Optional[int] = None
    space_before: Optional[float] = None
    space_after: Optional[float] = None
    line_spacing: Optional[float] = None
    line_spacing_rule: Optional[int] = None
    left_indent: Optional[float] = None
    right_indent: Optional[float] = None
    first_line_indent: Optional[float] = None
    keep_together: Optional[bool] = None
    keep_with_next: Optional[bool] = None
    page_break_before: Optional[bool] = None
    widow_control: Optional[bool] = None
    outline_level: Optional[int] = None


@dataclass
class LogicalBlock:
    text: str
    word_count: int
    paragraph_style: ParagraphStyle
    runs: List[RunStyle]
    block_type: str
    body_element_indices: List[int] = field(default_factory=list)

    def __repr__(self):
        preview = self.text[:80] + "..." if len(self.text) > 80 else self.text
        return f"LogicalBlock(type={self.block_type}, words={self.word_count}, text='{preview}')"


# ---------------------------------------------------------------------------
# Markdown helpers
# ---------------------------------------------------------------------------

_MD_INLINE_RE = re.compile(
    r'(\*\*\*.+?\*\*\*|(?<![\w])___.+?___(?![\w])'                 # bold+italic
    r'|\*\*.+?\*\*|(?<![\w])__.+?__(?![\w])'                         # bold
    r'|~~.+?~~'                                                      # strikethrough
    r'|==.+?=='                                                       # highlight
    r'|\+\+.+?\+\+'                                                  # underline
    r'|~[^~]+?~'                                                      # subscript
    r'|\^[^\^]+?\^'                                                   # superscript
    r'|\*[^*]+?\*|(?<![\w])_[^_]+?_(?![\w])'                         # italic (_ не внутри слов: snake_case)
    r'|`[^`]+?`'                                                      # inline code
    r'|\[[^\]]+?\]\([^)\s]+?\))',                                    # hyperlink [text](url)
    re.DOTALL,
)

_MD_HEADING_RE = re.compile(r'^(#{1,6})\s+(.*)$')
_MD_BLOCKQUOTE_RE = re.compile(r'^>\s?(.*)$')
_MD_HR_RE = re.compile(r'^(-{3,}|\*{3,}|_{3,})\s*$')
_MD_BULLET_RE = re.compile(r'^(\s*)([-*+•])\s+(.*)$')
_MD_NUMBERED_RE = re.compile(r'^(\s*)(\d+)[.)]\s+(.*)$')

_PLACEHOLDER_RE = re.compile(r'\{\{.*?\}\}')

_MD_TABLE_SEP_RE = re.compile(r'^\|?\s*[-:]+[-|\s:]*$')
_MD_TABLE_CELL_SPLIT_RE = re.compile(r'(?<!\\)\|')

_WORD_RE = re.compile(r'\w+')
# Символы, недопустимые в XML 1.0 (частый мусор в ответах LLM) — lxml на них падает
_XML_INVALID_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]')
_SAFE_URL_RE = re.compile(r'^(https?://|mailto:)', re.IGNORECASE)

_MD_CODE_FONT = 'Consolas'
_HYPERLINK_REL_TYPE = RT.HYPERLINK
_HYPERLINK_COLOR = '0563C1'
_W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
_XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
_NSMAP = {'w': _W_NS}

# Ширина области текста A4 с полями 2 см (twips) — для сетки создаваемых таблиц
_DEFAULT_TEXT_WIDTH_TWIPS = 9638

_JC_TO_ALIGN = {
    'left': WD_ALIGN_PARAGRAPH.LEFT, 'start': WD_ALIGN_PARAGRAPH.LEFT,
    'center': WD_ALIGN_PARAGRAPH.CENTER,
    'right': WD_ALIGN_PARAGRAPH.RIGHT, 'end': WD_ALIGN_PARAGRAPH.RIGHT,
    'both': WD_ALIGN_PARAGRAPH.JUSTIFY,
    'distribute': WD_ALIGN_PARAGRAPH.DISTRIBUTE,
}
_LINE_RULE_MAP = {
    'auto': WD_LINE_SPACING.MULTIPLE,
    'exact': WD_LINE_SPACING.EXACTLY,
    'atleast': WD_LINE_SPACING.AT_LEAST,
}
_HEADING_STYLE_MARKERS = ('heading', 'title', 'head', 'заголовок')
_HEADING_STYLE_EXCLUDE = ('header', 'footer', 'колонтитул')
_LIST_STYLE_MARKERS = ('list', 'bullet', 'number', 'список')


def _local_tag(elem) -> Optional[str]:
    """Локальное имя тега без namespace; None для комментариев/PI."""
    tag = elem.tag
    if not isinstance(tag, str):
        return None
    return tag.split('}', 1)[-1]


def _sanitize_xml_text(text) -> str:
    if text is None:
        return ''
    return _XML_INVALID_RE.sub('', str(text))


def _on_off(val) -> bool:
    """Значение атрибута w:val для on/off-свойств OOXML (отсутствие = true)."""
    if val is None:
        return True
    return str(val).lower() not in ('0', 'false', 'off')


def _to_int(val) -> Optional[int]:
    try:
        return int(float(val))
    except (TypeError, ValueError):
        return None

_RPR_ORDER = [
    'w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps',
    'w:smallCaps', 'w:strike', 'w:dstrike', 'w:outline', 'w:shadow',
    'w:emboss', 'w:imprint', 'w:noProof', 'w:snapToGrid', 'w:vanish',
    'w:webHidden', 'w:color', 'w:spacing', 'w:w', 'w:kern', 'w:position',
    'w:sz', 'w:szCs', 'w:highlight', 'w:u', 'w:effect', 'w:bdr', 'w:shd',
    'w:fitText', 'w:vertAlign', 'w:rtl', 'w:cs', 'w:em', 'w:lang',
    'w:eastAsianLayout', 'w:specVanish', 'w:oMath',
]

_PPR_ORDER = [
    'w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore',
    'w:framePr', 'w:widowControl', 'w:numPr', 'w:suppressLineNumbers',
    'w:pBdr', 'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku',
    'w:wordWrap', 'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE',
    'w:autoSpaceDN', 'w:bidi', 'w:adjustRightInd', 'w:snapToGrid',
    'w:spacing', 'w:ind', 'w:contextualSpacing', 'w:mirrorIndents',
    'w:suppressOverlap', 'w:jc', 'w:textDirection', 'w:textAlignment',
    'w:textboxTightWrap', 'w:outlineLvl', 'w:divId', 'w:cnfStyle',
    'w:rPr', 'w:sectPr', 'w:pPrChange',
]


class DocxWorker:
    def __init__(self, min_words: int = 10, merge_same_style: bool = True,
                 preserve_headings: bool = True, preserve_list_items: bool = False,
                 base_font_name: str = 'Times New Roman',
                 base_font_size: float = 14.0,
                 base_alignment=WD_ALIGN_PARAGRAPH.JUSTIFY):
        if not isinstance(min_words, int) or isinstance(min_words, bool) or min_words < 0:
            raise ValueError(f'min_words \u0434\u043E\u043B\u0436\u0435\u043D \u0431\u044B\u0442\u044C \u0446\u0435\u043B\u044B\u043C \u0447\u0438\u0441\u043B\u043E\u043C >= 0, \u043F\u043E\u043B\u0443\u0447\u0435\u043D\u043E: {min_words!r}')
        try:
            base_font_size = float(base_font_size)
        except (TypeError, ValueError):
            raise ValueError(f'base_font_size \u0434\u043E\u043B\u0436\u0435\u043D \u0431\u044B\u0442\u044C \u0447\u0438\u0441\u043B\u043E\u043C, \u043F\u043E\u043B\u0443\u0447\u0435\u043D\u043E: {base_font_size!r}')
        if base_font_size <= 0:
            raise ValueError(f'base_font_size \u0434\u043E\u043B\u0436\u0435\u043D \u0431\u044B\u0442\u044C > 0, \u043F\u043E\u043B\u0443\u0447\u0435\u043D\u043E: {base_font_size}')
        self.min_words = min_words
        self.base_font_name = base_font_name
        self.base_font_size = base_font_size
        self.base_alignment = base_alignment
        self.merge_same_style = merge_same_style
        self.preserve_headings = preserve_headings
        self.preserve_list_items = preserve_list_items
        self._doc = None
        self._numbering = None
        self._next_num_id = 1
        self._bullet_num_id = None
        self._decimal_abstract_id = None
        self._cur_decimal = None
        self._prev_is_list = False
        self._style_names: Dict[str, str] = {}

    @staticmethod
    def _validate_docx_path(path: str, name: str = 'input_path') -> None:
        if not path or not isinstance(path, (str, os.PathLike)):
            raise ValueError(f'{name}: \u043D\u0435 \u0437\u0430\u0434\u0430\u043D \u043F\u0443\u0442\u044C \u043A \u0444\u0430\u0439\u043B\u0443')
        if not os.path.isfile(path):
            raise FileNotFoundError(f'{name}: \u0444\u0430\u0439\u043B \u043D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D: {path}')
        if not zipfile.is_zipfile(path):
            raise ValueError(f'{name}: \u0444\u0430\u0439\u043B \u043D\u0435 \u044F\u0432\u043B\u044F\u0435\u0442\u0441\u044F \u0434\u043E\u043A\u0443\u043C\u0435\u043D\u0442\u043E\u043C .docx (OOXML): {path}')

    @staticmethod
    def _ensure_parent_dir(path: str) -> None:
        parent = os.path.dirname(os.path.abspath(path))
        if parent:
            os.makedirs(parent, exist_ok=True)

    def _load_style_names(self, doc) -> None:
        """styleId -> \u0438\u043C\u044F \u0441\u0442\u0438\u043B\u044F \u0432 \u043D\u0438\u0436\u043D\u0435\u043C \u0440\u0435\u0433\u0438\u0441\u0442\u0440\u0435. \u0412 \u0440\u0443\u0441\u0441\u043A\u043E\u044F\u0437\u044B\u0447\u043D\u044B\u0445 \u0448\u0430\u0431\u043B\u043E\u043D\u0430\u0445
        styleId \u0437\u0430\u0433\u043E\u043B\u043E\u0432\u043A\u043E\u0432 \u0447\u0430\u0441\u0442\u043E '1', '2', ..., \u0430 \u0438\u043C\u044F \u2014 'heading 1'."""
        names = {}
        try:
            for st in doc.styles.element.findall(qn('w:style')):
                sid = st.get(qn('w:styleId'))
                name_el = st.find(qn('w:name'))
                if sid and name_el is not None and name_el.get(qn('w:val')):
                    names[sid] = name_el.get(qn('w:val')).lower()
        except Exception as e:
            logger.warning(f'DocxWorker: \u043D\u0435 \u0443\u0434\u0430\u043B\u043E\u0441\u044C \u043F\u0440\u043E\u0447\u0438\u0442\u0430\u0442\u044C \u0441\u0442\u0438\u043B\u0438 \u0434\u043E\u043A\u0443\u043C\u0435\u043D\u0442\u0430: {e}')
        self._style_names = names

    def _style_matches(self, style_id: Optional[str], markers, exclude=()) -> bool:
        if not style_id:
            return False
        for candidate in (style_id.lower(), self._style_names.get(style_id, '')):
            if not candidate or any(x in candidate for x in exclude):
                continue
            if any(x in candidate for x in markers):
                return True
        return False

    def _count_words(self, text: str) -> int:
        if not text or not text.strip():
            return 0
        # \w \u0432 Python 3 \u0443\u0436\u0435 \u0432\u043A\u043B\u044E\u0447\u0430\u0435\u0442 \u043A\u0438\u0440\u0438\u043B\u043B\u0438\u0446\u0443
        return len(_WORD_RE.findall(text))

    def _extract_run_style(self, run) -> RunStyle:
        rs = RunStyle(text=run.text)
        rs.bold = run.bold
        rs.italic = run.italic
        rs.underline = run.underline
        rs.font_name = run.font.name
        rs.font_size = run.font.size.pt if run.font.size else None
        try:
            rgb = run.font.color.rgb if run.font.color is not None else None
            if rgb is not None:
                rs.font_color = (rgb[0], rgb[1], rgb[2])
        except Exception:
            pass
        rs.all_caps = run.font.all_caps
        rs.small_caps = run.font.small_caps
        rs.strike = run.font.strike
        rs.double_strike = run.font.double_strike
        rs.subscript = run.font.subscript
        rs.superscript = run.font.superscript
        rs.shadow = run.font.shadow
        rs.outline = run.font.outline
        rs.emboss = run.font.emboss
        rs.imprint = run.font.imprint
        if run.style and run.style.name:
            rs.character_style = run.style.name
        try:
            rPr = run._r.find(qn('w:rPr'))
            if rPr is not None:
                highlight = rPr.find(qn('w:highlight'))
                if highlight is not None:
                    val = highlight.get(qn('w:val'))
                    if val:
                        # brightGreen -> BRIGHT_GREEN
                        name = re.sub(r'([a-z])([A-Z])', r'_', val).upper()
                        rs.highlight_color = getattr(WD_COLOR_INDEX, name, None)
        except Exception:
            pass
        return rs

    def _extract_paragraph_style(self, paragraph) -> ParagraphStyle:
        ps = ParagraphStyle()
        if paragraph.style and paragraph.style.name:
            ps.style_name = paragraph.style.name
        pf = paragraph.paragraph_format
        ps.alignment = pf.alignment
        ps.space_before = pf.space_before.pt if pf.space_before else None
        ps.space_after = pf.space_after.pt if pf.space_after else None
        ps.line_spacing = pf.line_spacing
        ps.line_spacing_rule = pf.line_spacing_rule
        ps.left_indent = pf.left_indent.pt if pf.left_indent else None
        ps.right_indent = pf.right_indent.pt if pf.right_indent else None
        ps.first_line_indent = pf.first_line_indent.pt if pf.first_line_indent else None
        ps.keep_together = pf.keep_together
        ps.keep_with_next = pf.keep_with_next
        ps.page_break_before = pf.page_break_before
        ps.widow_control = pf.widow_control
        try:
            pPr = paragraph._p.find(qn('w:pPr'))
            if pPr is not None:
                outline = pPr.find(qn('w:outlineLvl'))
                if outline is not None:
                    val = outline.get(qn('w:val'))
                    if val is not None:
                        ps.outline_level = _to_int(val)
        except Exception:
            pass
        return ps

    def _styles_compatible(self, style1: ParagraphStyle, style2: ParagraphStyle) -> bool:
        if style1.style_name != style2.style_name:
            return False
        if style1.alignment != style2.alignment:
            return False
        def approx_equal(a, b, tolerance=0.5):
            if a is None and b is None:
                return True
            if a is None or b is None:
                return False
            return abs(a - b) < tolerance
        if not approx_equal(style1.left_indent, style2.left_indent):
            return False
        if not approx_equal(style1.right_indent, style2.right_indent):
            return False
        if not approx_equal(style1.first_line_indent, style2.first_line_indent):
            return False
        if not approx_equal(style1.space_before, style2.space_before):
            return False
        if not approx_equal(style1.space_after, style2.space_after):
            return False
        return True

    def _is_heading(self, paragraph) -> bool:
        if paragraph.style and paragraph.style.name:
            style_name = paragraph.style.name.lower()
            if (any(x in style_name for x in _HEADING_STYLE_MARKERS)
                    and not any(x in style_name for x in _HEADING_STYLE_EXCLUDE)):
                return True
        return self._has_heading_outline(paragraph._p)

    def _is_list_item(self, paragraph) -> bool:
        if paragraph.style and paragraph.style.name:
            style_name = paragraph.style.name.lower()
            if any(x in style_name for x in _LIST_STYLE_MARKERS):
                return True
        pPr = paragraph._p.find(qn('w:pPr'))
        return pPr is not None and pPr.find(qn('w:numPr')) is not None

    @staticmethod
    def _has_heading_outline(p_elem) -> bool:
        """outlineLvl 0..8 — уровень заголовка; 9 — обычный текст."""
        pPr = p_elem.find(qn('w:pPr'))
        if pPr is None:
            return False
        outline = pPr.find(qn('w:outlineLvl'))
        if outline is None:
            return False
        lvl = _to_int(outline.get(qn('w:val')))
        return lvl is not None and 0 <= lvl < 9

    def _extract_paragraph_style_from_xml(self, p_elem) -> ParagraphStyle:
        ps = ParagraphStyle()
        try:
            pPr = p_elem.find(qn('w:pPr'))
            if pPr is None:
                return ps
            pStyle = pPr.find(qn('w:pStyle'))
            if pStyle is not None:
                val = pStyle.get(qn('w:val'))
                if val:
                    ps.style_name = val
            jc = pPr.find(qn('w:jc'))
            if jc is not None:
                val = jc.get(qn('w:val'))
                if val:
                    ps.alignment = _JC_TO_ALIGN.get(val.lower())
            spacing = pPr.find(qn('w:spacing'))
            if spacing is not None:
                before = _to_int(spacing.get(qn('w:before')))
                after = _to_int(spacing.get(qn('w:after')))
                line = _to_int(spacing.get(qn('w:line')))
                line_rule = (spacing.get(qn('w:lineRule')) or 'auto').lower()
                if before is not None:
                    ps.space_before = before / 20.0
                if after is not None:
                    ps.space_after = after / 20.0
                ps.line_spacing_rule = _LINE_RULE_MAP.get(line_rule)
                if line is not None:
                    # auto: в 240-х долях строки (множитель),
                    # exact/atLeast: в twips (абсолютное значение)
                    ps.line_spacing = line / 240.0 if line_rule == 'auto' else Pt(line / 20.0)
            ind = pPr.find(qn('w:ind'))
            if ind is not None:
                left = _to_int(ind.get(qn('w:left')) or ind.get(qn('w:start')))
                right = _to_int(ind.get(qn('w:right')) or ind.get(qn('w:end')))
                first_line = _to_int(ind.get(qn('w:firstLine')))
                hanging = _to_int(ind.get(qn('w:hanging')))
                if left is not None:
                    ps.left_indent = left / 20.0
                if right is not None:
                    ps.right_indent = right / 20.0
                if first_line is not None:
                    ps.first_line_indent = first_line / 20.0
                elif hanging is not None:
                    ps.first_line_indent = -hanging / 20.0
            outline = pPr.find(qn('w:outlineLvl'))
            if outline is not None:
                ps.outline_level = _to_int(outline.get(qn('w:val')))
            for prop_name, attr_name in (
                ('w:keepLines', 'keep_together'),
                ('w:keepNext', 'keep_with_next'),
                ('w:pageBreakBefore', 'page_break_before'),
                ('w:widowControl', 'widow_control'),
            ):
                prop = pPr.find(qn(prop_name))
                if prop is not None:
                    setattr(ps, attr_name, _on_off(prop.get(qn('w:val'))))
        except Exception as e:
            logger.debug(f'DocxWorker: не удалось разобрать стиль абзаца: {e}')
        return ps

    def _is_heading_from_xml(self, p_elem) -> bool:
        if self._has_heading_outline(p_elem):
            return True
        pPr = p_elem.find(qn('w:pPr'))
        if pPr is None:
            return False
        pStyle = pPr.find(qn('w:pStyle'))
        if pStyle is None:
            return False
        return self._style_matches(pStyle.get(qn('w:val')),
                                   _HEADING_STYLE_MARKERS, _HEADING_STYLE_EXCLUDE)

    def _is_list_item_from_xml(self, p_elem) -> bool:
        pPr = p_elem.find(qn('w:pPr'))
        if pPr is None:
            return False
        if pPr.find(qn('w:numPr')) is not None:
            return True
        pStyle = pPr.find(qn('w:pStyle'))
        if pStyle is None:
            return False
        return self._style_matches(pStyle.get(qn('w:val')), _LIST_STYLE_MARKERS)

    def _apply_paragraph_style(self, new_para, style: ParagraphStyle):
        if style.style_name:
            try:
                new_para.style = style.style_name
            except Exception:
                logger.debug(f'DocxWorker: стиль "{style.style_name}" не найден в документе')
        pf = new_para.paragraph_format
        if style.alignment is not None:
            pf.alignment = style.alignment
        if style.space_before is not None:
            pf.space_before = Pt(style.space_before)
        if style.space_after is not None:
            pf.space_after = Pt(style.space_after)
        if style.line_spacing is not None:
            pf.line_spacing = style.line_spacing
        if style.line_spacing_rule is not None:
            pf.line_spacing_rule = style.line_spacing_rule
        if style.left_indent is not None:
            pf.left_indent = Pt(style.left_indent)
        if style.right_indent is not None:
            pf.right_indent = Pt(style.right_indent)
        if style.first_line_indent is not None:
            pf.first_line_indent = Pt(style.first_line_indent)
        if style.keep_together is not None:
            pf.keep_together = style.keep_together
        if style.keep_with_next is not None:
            pf.keep_with_next = style.keep_with_next
        if style.page_break_before is not None:
            pf.page_break_before = style.page_break_before
        if style.widow_control is not None:
            pf.widow_control = style.widow_control

    # ------------------------------------------------------------------
    # Block building
    # ------------------------------------------------------------------

    def _paragraph_text(self, p_elem, nsmap) -> str:
        return ''.join(t.text or '' for t in p_elem.iterfind('.//w:t', namespaces=nsmap)).strip()

    def _build_raw_blocks(self, doc, nsmap) -> List[Dict[str, Any]]:
        self._load_style_names(doc)
        raw_blocks = []
        body = doc.element.body
        for idx, element in enumerate(list(body)):
            tag = _local_tag(element)
            if tag == 'tbl':
                text = ''.join(
                    t.text for t in element.iterfind('.//w:t', namespaces=nsmap) if t.text
                ).strip()
                raw_blocks.append({
                    'text': text,
                    'word_count': self._count_words(text),
                    'para_style': ParagraphStyle(),
                    'runs': [],
                    'type': 'table',
                    'element_index': idx,
                    'element': element,
                    'is_table': True,
                })
            elif tag == 'p':
                text = self._paragraph_text(element, nsmap)
                para_style = self._extract_paragraph_style_from_xml(element)
                if self._is_heading_from_xml(element):
                    block_type = 'heading'
                elif self._is_list_item_from_xml(element):
                    block_type = 'list_item'
                else:
                    block_type = 'paragraph'
                raw_blocks.append({
                    'text': text,
                    'word_count': self._count_words(text),
                    'para_style': para_style,
                    'runs': [],
                    'type': block_type,
                    'element_index': idx,
                    'element': element,
                    'is_table': False,
                    'has_placeholder': bool(_PLACEHOLDER_RE.search(text)),
                })
        return raw_blocks

    def _merge_raw_blocks(self, raw_blocks: List[Dict[str, Any]],
                          preserve_headings: bool,
                          preserve_list_items: bool) -> List[Dict[str, Any]]:
        blocks = []
        i = 0
        while i < len(raw_blocks):
            current = raw_blocks[i]
            is_atomic = (
                current['is_table']
                or current.get('has_placeholder')
                or (preserve_headings and current['type'] == 'heading' and current['word_count'] > 0)
                or (preserve_list_items and current['type'] == 'list_item' and current['word_count'] > 0)
            )
            if is_atomic:
                blocks.append({
                    'blocks': [current],
                    'is_table': current['is_table'],
                    'type': current['type'],
                    'text': current['text'],
                    'word_count': current['word_count'],
                })
                i += 1
                continue
            group = [current]
            merged_text = current['text']
            merged_words = current['word_count']
            j = i + 1
            while merged_words < self.min_words and j < len(raw_blocks):
                nxt = raw_blocks[j]
                if nxt['is_table'] or nxt.get('has_placeholder'):
                    break
                if preserve_headings and nxt['type'] == 'heading' and nxt['word_count'] > 0:
                    break
                if preserve_list_items and nxt['type'] == 'list_item' and nxt['word_count'] > 0:
                    break
                group.append(nxt)
                merged_words += nxt['word_count']
                if merged_text:
                    merged_text += '\n'
                merged_text += nxt['text']
                j += 1
            first_type = group[0]['type']
            merged_type = first_type if all(b['type'] == first_type for b in group) else 'merged'
            blocks.append({
                'blocks': group,
                'is_table': False,
                'type': merged_type,
                'text': merged_text,
                'word_count': merged_words,
            })
            i = j
        return blocks

    def split(self, docx_path: str) -> List[LogicalBlock]:
        self._validate_docx_path(docx_path, 'docx_path')
        doc = Document(docx_path)
        nsmap = _NSMAP
        raw_blocks = self._build_raw_blocks(doc, nsmap)
        merged = self._merge_raw_blocks(raw_blocks, self.preserve_headings, self.preserve_list_items)
        # Индекс элемент -> Paragraph строится один раз (раньше — полный
        # перебор doc.paragraphs на каждый блок, O(N^2))
        para_by_elem = {p._p: p for p in doc.paragraphs}
        blocks = []
        for m in merged:
            if m['is_table']:
                runs: List[RunStyle] = []
                para_style = ParagraphStyle()
            else:
                para_objs = [para_by_elem[b['element']] for b in m['blocks']
                             if b['element'] in para_by_elem]
                runs = []
                for p in para_objs:
                    runs.extend(self._extract_run_style(run) for run in p.runs)
                para_style = self._extract_paragraph_style(para_objs[0]) if para_objs else ParagraphStyle()
            blocks.append(LogicalBlock(
                text=m['text'],
                word_count=m['word_count'],
                paragraph_style=para_style,
                runs=runs,
                block_type=m['type'],
                body_element_indices=[b['element_index'] for b in m['blocks']],
            ))
        return blocks

    def _copy_docx_with_rels(self, source_path: str, dest_path: str) -> None:
        if os.path.abspath(source_path) == os.path.abspath(dest_path):
            return
        self._ensure_parent_dir(dest_path)
        shutil.copy2(source_path, dest_path)

    def _save_block_from_source(self, source, block: LogicalBlock, output_path: str) -> None:
        """source — путь или BytesIO с содержимым исходного .docx."""
        doc = Document(source)
        body = doc.element.body
        keep_indices = set(block.body_element_indices)
        all_elements = list(body)
        if all_elements and _local_tag(all_elements[-1]) == 'sectPr':
            keep_indices.add(len(all_elements) - 1)
        for idx in range(len(all_elements) - 1, -1, -1):
            if idx not in keep_indices:
                body.remove(all_elements[idx])
        self._validate_document(doc)
        self._ensure_parent_dir(output_path)
        doc.save(output_path)

    def create_block_document(self, block: LogicalBlock, output_path: str,
                               source_docx_path: str = None) -> None:
        if not isinstance(block, LogicalBlock):
            raise TypeError('create_block_document: ожидается LogicalBlock')
        if not output_path:
            raise ValueError('create_block_document: не задан output_path')
        if source_docx_path:
            self._validate_docx_path(source_docx_path, 'source_docx_path')
            self._save_block_from_source(source_docx_path, block, output_path)
        else:
            # Без исходника переносить нечего — создаётся пустой документ
            logger.warning('create_block_document: source_docx_path не задан, создан пустой документ')
            self._ensure_parent_dir(output_path)
            Document().save(output_path)

    def split_to_files(self, docx_path: str, output_prefix: str) -> List[str]:
        blocks = self.split(docx_path)
        # Исходник читается с диска один раз, далее каждый блок собирается из памяти
        with open(docx_path, 'rb') as f:
            src_bytes = f.read()
        output_paths = []
        for idx, block in enumerate(blocks):
            output_path = f"{output_prefix}p{idx}d-{text_random(10)}.docx"
            self._save_block_from_source(BytesIO(src_bytes), block, output_path)
            output_paths.append(output_path)
        return output_paths

    # ------------------------------------------------------------------
    # Markdown -> Word conversion helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _parse_markdown_inline(text: str) -> List[Tuple[str, Dict[str, Any]]]:
        segments: List[Tuple[str, Dict[str, Any]]] = []
        pos = 0
        for m in _MD_INLINE_RE.finditer(text):
            if m.start() > pos:
                segments.append((text[pos:m.start()], {}))
            token = m.group(0)
            if token.startswith('***') and token.endswith('***') and len(token) > 6:
                segments.append((token[3:-3], {'bold': True, 'italic': True}))
            elif token.startswith('___') and token.endswith('___') and len(token) > 6:
                segments.append((token[3:-3], {'bold': True, 'italic': True}))
            elif token.startswith('**') and token.endswith('**') and len(token) > 4:
                segments.append((token[2:-2], {'bold': True}))
            elif token.startswith('__') and token.endswith('__') and len(token) > 4:
                segments.append((token[2:-2], {'bold': True}))
            elif token.startswith('~~') and token.endswith('~~') and len(token) > 4:
                segments.append((token[2:-2], {'strike': True}))
            elif token.startswith('==') and token.endswith('==') and len(token) > 4:
                segments.append((token[2:-2], {'highlight': True}))
            elif token.startswith('++') and token.endswith('++') and len(token) > 4:
                segments.append((token[2:-2], {'underline': True}))
            elif token.startswith('~') and token.endswith('~') and len(token) > 2 and not token.startswith('~~'):
                segments.append((token[1:-1], {'subscript': True}))
            elif token.startswith('^') and token.endswith('^') and len(token) > 2:
                segments.append((token[1:-1], {'superscript': True}))
            elif token.startswith('`') and token.endswith('`') and len(token) > 2:
                segments.append((token[1:-1], {'code': True}))
            elif token.startswith('[') and token.endswith(')') and '](' in token:
                inner = token[1:-1]
                bracket_idx = inner.index('](')
                link_text = inner[:bracket_idx]
                link_url = inner[bracket_idx + 2:]
                segments.append((link_text, {'hyperlink': link_url}))
            elif token.startswith('*') and token.endswith('*') and len(token) > 2:
                segments.append((token[1:-1], {'italic': True}))
            elif token.startswith('_') and token.endswith('_') and len(token) > 2:
                segments.append((token[1:-1], {'italic': True}))
            else:
                segments.append((token, {}))
            pos = m.end()
        if pos < len(text):
            segments.append((text[pos:], {}))
        expanded: List[Tuple[str, Dict[str, Any]]] = []
        for chunk, style in segments:
            # содержимое inline-кода не разбирается как markdown
            if not style or style.get('code'):
                expanded.append((chunk, style))
                continue
            inner = DocxWorker._parse_markdown_inline(chunk)
            for inner_chunk, inner_style in inner:
                merged_style = dict(style)
                merged_style.update(inner_style)
                expanded.append((inner_chunk, merged_style))
        merged: List[Tuple[str, Dict[str, Any]]] = []
        for chunk, style in expanded:
            if merged and merged[-1][1] == style:
                merged[-1] = (merged[-1][0] + chunk, style)
            else:
                merged.append((chunk, style))
        return merged

    @staticmethod
    def _get_or_add_rPr(r_elem):
        rPr = r_elem.find(qn('w:rPr'))
        if rPr is None:
            rPr = OxmlElement('w:rPr')
            r_elem.insert(0, rPr)
        return rPr

    @classmethod
    def _set_run_bool_prop(cls, rPr, tag: str, on: bool = True) -> None:
        el = rPr.find(qn(tag))
        if on:
            if el is None:
                el = OxmlElement(tag)
                cls._insert_in_order(rPr, el, _RPR_ORDER)
            el.set(qn('w:val'), '1')
        elif el is not None:
            rPr.remove(el)

    @staticmethod
    def _insert_in_order(parent, child, order_list):
        child_tag = child.tag
        child_idx = None
        for i, tag_name in enumerate(order_list):
            if qn(tag_name) == child_tag:
                child_idx = i
                break
        if child_idx is None:
            parent.append(child)
            return
        for existing in parent:
            existing_idx = None
            for i, tag_name in enumerate(order_list):
                if qn(tag_name) == existing.tag:
                    existing_idx = i
                    break
            if existing_idx is None:
                continue
            if existing_idx > child_idx:
                existing.addprevious(child)
                return
        parent.append(child)

    def _make_run(self, text: str, style: Dict[str, Any], template_rPr) -> Any:
        r = OxmlElement('w:r')
        rPr = deepcopy(template_rPr) if template_rPr is not None else OxmlElement('w:rPr')

        # Boolean text properties
        self._set_run_bool_prop(rPr, 'w:b', style.get('bold', False))
        self._set_run_bool_prop(rPr, 'w:i', style.get('italic', False))
        self._set_run_bool_prop(rPr, 'w:strike', style.get('strike', False))

        # Underline
        if style.get('underline'):
            u = rPr.find(qn('w:u'))
            if u is None:
                u = OxmlElement('w:u')
                self._insert_in_order(rPr, u, _RPR_ORDER)
            u.set(qn('w:val'), 'single')

        # Highlight
        if style.get('highlight'):
            hl = rPr.find(qn('w:highlight'))
            if hl is None:
                hl = OxmlElement('w:highlight')
                self._insert_in_order(rPr, hl, _RPR_ORDER)
            hl.set(qn('w:val'), 'yellow')

        # Subscript
        if style.get('subscript'):
            va = rPr.find(qn('w:vertAlign'))
            if va is None:
                va = OxmlElement('w:vertAlign')
                self._insert_in_order(rPr, va, _RPR_ORDER)
            va.set(qn('w:val'), 'subscript')

        # Superscript
        if style.get('superscript'):
            va = rPr.find(qn('w:vertAlign'))
            if va is None:
                va = OxmlElement('w:vertAlign')
                self._insert_in_order(rPr, va, _RPR_ORDER)
            va.set(qn('w:val'), 'superscript')

        # FIX v7: Only set font/size if template_rPr already has them.
        # This preserves the etalon's clean style where runs don't have
        # explicit font/size unless the original template had them.
        if style.get('code'):
            font_name = _MD_CODE_FONT
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                self._insert_in_order(rPr, rFonts, _RPR_ORDER)
            for attr in ('w:ascii', 'w:hAnsi', 'w:cs'):
                rFonts.set(qn(attr), font_name)
        # Do NOT add font/size if template didn't have them — let document defaults apply

        r.append(rPr)

        text = _sanitize_xml_text(text)
        # 	 внутри w:t Word не отображает как табуляцию — нужен w:tab
        for i, piece in enumerate(text.split('	')):
            if i:
                r.append(OxmlElement('w:tab'))
            if piece:
                t = OxmlElement('w:t')
                t.text = piece
                if piece != piece.strip():
                    t.set(_XML_SPACE, 'preserve')
                r.append(t)
        return r

    def _make_hyperlink(self, text: str, url: str, style: Dict[str, Any], template_rPr) -> Any:
        """w:hyperlink с внешней ссылкой; без документа или для небезопасной
        схемы URL возвращает обычный run с текстом ссылки."""
        url = (url or '').strip()
        if self._doc is None or not _SAFE_URL_RE.match(url):
            return self._make_run(text, style, template_rPr)
        try:
            r_id = self._doc.part.relate_to(url, _HYPERLINK_REL_TYPE, is_external=True)
        except Exception as e:
            logger.warning(f'DocxWorker: не удалось создать гиперссылку {url!r}: {e}')
            return self._make_run(text, style, template_rPr)
        link_style = dict(style)
        link_style.pop('hyperlink', None)
        link_style['underline'] = True
        run = self._make_run(text, link_style, template_rPr)
        rPr = run.find(qn('w:rPr'))
        color = rPr.find(qn('w:color'))
        if color is None:
            color = OxmlElement('w:color')
            self._insert_in_order(rPr, color, _RPR_ORDER)
        color.set(qn('w:val'), _HYPERLINK_COLOR)
        hyperlink = OxmlElement('w:hyperlink')
        hyperlink.set(qn('r:id'), r_id)
        hyperlink.set(qn('w:history'), '1')
        hyperlink.append(run)
        return hyperlink

    def _append_inline(self, p_elem, text: str, template_rPr, force_bold: bool = False) -> None:
        for chunk, style in self._parse_markdown_inline(text):
            if chunk == '':
                continue
            if force_bold and not style.get('bold'):
                style = dict(style, bold=True)
            if style.get('hyperlink'):
                p_elem.append(self._make_hyperlink(chunk, style['hyperlink'], style, template_rPr))
            else:
                p_elem.append(self._make_run(chunk, style, template_rPr))

    @staticmethod
    def _strip_unwanted_pPr_elements(p_elem) -> None:
        pPr = p_elem.find(qn('w:pPr'))
        if pPr is not None:
            for tag_name in ('w:sectPr', 'w:numPr'):
                el = pPr.find(qn(tag_name))
                if el is not None:
                    pPr.remove(el)

    def _clear_paragraph_runs(self, p_elem, nsmap) -> Any:
        """Remove all run-level content, return a template w:rPr (or None).
        FIX v7: Preserve runs that contain only w:tab/w:br (no w:t text)
        so that tab characters from the original template are kept."""
        template_rPr = None
        for r in p_elem.findall('w:r', namespaces=nsmap):
            # Check if this run has any w:t (text) element
            t_elem = r.find(qn('w:t'))
            if t_elem is None:
                # This run has no text content (e.g., w:tab, w:br) — keep it
                continue
            if template_rPr is None:
                rPr = r.find(qn('w:rPr'))
                if rPr is not None:
                    template_rPr = deepcopy(rPr)
            p_elem.remove(r)

        run_level_tags = (
            'w:hyperlink', 'w:ins', 'w:del', 'w:smartTag',
            'w:bookmarkStart', 'w:bookmarkEnd',
            'w:commentRangeStart', 'w:commentRangeEnd', 'w:commentReference',
            'w:proofErr', 'w:permStart', 'w:permEnd',
            'w:moveFrom', 'w:moveTo',
            'w:moveFromRangeStart', 'w:moveFromRangeEnd',
            'w:moveToRangeStart', 'w:moveToRangeEnd',
            'w:customXml', 'w:fldSimple', 'w:fldChar',
            'w:object', 'w:pict', 'w:txbxContent', 'w:sdt',
        )
        for tag in run_level_tags:
            for el in p_elem.findall(tag, namespaces=nsmap):
                if template_rPr is None:
                    r = el.find('.//w:r', namespaces=nsmap)
                    if r is not None:
                        rPr = r.find(qn('w:rPr'))
                        if rPr is not None:
                            template_rPr = deepcopy(rPr)
                p_elem.remove(el)
        return template_rPr

    def _apply_markdown_to_paragraph(self, p_elem, markdown_text: str, nsmap) -> None:
        """Replace paragraph content with markdown-rendered runs."""
        text = markdown_text
        text = text.replace('\r\n', '\n').replace('\r', '\n')
        text = text.replace('\\n', '\n')

        stripped = text.strip() if text else ''

        # --- Heading detection ---
        heading_level = 0
        hm = _MD_HEADING_RE.match(stripped) if stripped else None
        if hm:
            heading_level = len(hm.group(1))
            text = hm.group(2)
            stripped = text.strip()

        template_rPr = self._clear_paragraph_runs(p_elem, nsmap)

        if heading_level:
            self._apply_heading_style(p_elem, heading_level)

        # FIX v8: Do NOT apply base alignment. The template's alignment
        # is already correct and matches the etalon. Applying base_alignment
        # (JUSTIFY) would overwrite jc=right and add jc=both where there
        # should be no jc at all (e.g., "Рапорт" with jc=None).

        if not text or not text.strip():
            self._prev_is_list = False
            return

        # --- Blockquote ---
        bq = _MD_BLOCKQUOTE_RE.match(stripped)
        if bq:
            text = bq.group(1)
            self._apply_blockquote_style(p_elem)

        # --- Horizontal rule ---
        if _MD_HR_RE.match(stripped):
            self._apply_page_break(p_elem)
            return

        # --- Bullet / numbered list ---
        # FIX v13: ids come from numbering definitions we own (they never
        # clash with the template's numbering.xml); every new numbered list
        # restarts from 1; alignment is inherited from the paragraph itself.
        line_text = text.rstrip().replace('	', '  ')
        bullet_m = _MD_BULLET_RE.match(line_text)
        num_m = None if bullet_m else _MD_NUMBERED_RE.match(line_text)
        if (bullet_m or num_m) and self._numbering is None:
            # Нумерация не подготовлена (_ensure_numbering_definitions не вызывался):
            # маркер списка остаётся текстом, а не ссылкой на несуществующий numId
            bullet_m = num_m = None
        if bullet_m:
            level = len(bullet_m.group(1)) // 2
            text = bullet_m.group(3)
            self._apply_list_style(p_elem, 'bullet', level, self._bullet_num_id)
            self._prev_is_list = True
        elif num_m:
            level = len(num_m.group(1)) // 2
            text = num_m.group(3)
            if not self._prev_is_list or self._cur_decimal is None:
                self._cur_decimal = self._new_decimal_num()
            self._apply_list_style(p_elem, 'number', level, self._cur_decimal)
            self._prev_is_list = True
        else:
            self._prev_is_list = False

        self._append_inline(p_elem, text, template_rPr)

    @classmethod
    def _apply_blockquote_style(cls, p_elem) -> None:
        pPr = p_elem.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            p_elem.insert(0, pPr)
        ind = pPr.find(qn('w:ind'))
        if ind is None:
            ind = OxmlElement('w:ind')
            cls._insert_in_order(pPr, ind, _PPR_ORDER)
        ind.set(qn('w:left'), '708')
        ind.set(qn('w:hanging'), '354')
        p_rPr = pPr.find(qn('w:rPr'))
        if p_rPr is None:
            p_rPr = OxmlElement('w:rPr')
            cls._insert_in_order(pPr, p_rPr, _PPR_ORDER)
        i_elem = p_rPr.find(qn('w:i'))
        if i_elem is None:
            i_elem = OxmlElement('w:i')
            cls._insert_in_order(p_rPr, i_elem, _RPR_ORDER)
        i_elem.set(qn('w:val'), '1')

    @staticmethod
    def _apply_page_break(p_elem) -> None:
        pPr = p_elem.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            p_elem.insert(0, pPr)
        pbb = pPr.find(qn('w:pageBreakBefore'))
        if pbb is None:
            pbb = OxmlElement('w:pageBreakBefore')
            pPr.insert(0, pbb)
        pbb.set(qn('w:val'), '1')

    @classmethod
    def _apply_list_style(cls, p_elem, bullet_type: str = 'bullet', level: int = 0,
                          num_id: int = 1) -> None:
        pPr = p_elem.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            p_elem.insert(0, pPr)
        for tag in ('w:ind', 'w:numPr'):
            el = pPr.find(qn(tag))
            if el is not None:
                pPr.remove(el)
        numPr = OxmlElement('w:numPr')
        cls._insert_in_order(pPr, numPr, _PPR_ORDER)
        ilvl = OxmlElement('w:ilvl')
        ilvl.set(qn('w:val'), str(min(level, 8)))
        numPr.append(ilvl)
        numId = OxmlElement('w:numId')
        numId.set(qn('w:val'), str(num_id))
        numPr.append(numId)
        # jc is NOT touched: the paragraph keeps the alignment of the template

    def _ensure_numbering_definitions(self, doc) -> None:
        """FIX v13: works with or without an existing numbering part. Adds our own
        abstractNum/num with ids above everything already in the template."""
        from docx.oxml import parse_xml
        from docx.opc.packuri import PackURI
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        from docx.parts.numbering import NumberingPart
        self._doc = doc
        try:
            numbering = doc.part.part_related_by(RT.NUMBERING).element
        except KeyError:
            numbering = parse_xml(f'<w:numbering xmlns:w="{_W_NS}"/>')
            part = NumberingPart(PackURI('/word/numbering.xml'),
                'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
                numbering, doc.part.package)
            doc.part.relate_to(part, RT.NUMBERING)
        self._numbering = numbering
        ab_ids = [i for i in (_to_int(e.get(qn('w:abstractNumId')))
                              for e in numbering.findall(qn('w:abstractNum'))) if i is not None]
        n_ids = [i for i in (_to_int(e.get(qn('w:numId')))
                             for e in numbering.findall(qn('w:num'))) if i is not None]
        bullet_ab = max(ab_ids + [-1]) + 1
        decimal_ab = bullet_ab + 1
        self._next_num_id = max(n_ids + [0]) + 1

        def lvl(i, fmt, text, left, font=None):
            rpr = (f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:hint="default"/></w:rPr>'
                   if font else '')
            return (f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
                    f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
                    f'<w:pPr><w:ind w:left="{left}" w:hanging="360"/></w:pPr>{rpr}</w:lvl>')

        bullets = [('\u2022', None), ('\u2013', None), ('\u2022', None)]
        bullet_lvls = ''.join(lvl(i, 'bullet', bullets[i % 3][0], 720 + 720 * i) for i in range(9))
        dec_fmt = [('decimal', '%1.'), ('lowerLetter', '%2)'), ('lowerRoman', '%3.')]
        dec_lvls = ''.join(
            lvl(i, dec_fmt[i % 3][0], f'%{i + 1}' + (')' if i % 3 == 1 else '.'), 720 + 720 * i)
            for i in range(9))
        ab_b = parse_xml(f'<w:abstractNum xmlns:w="{_W_NS}" w:abstractNumId="{bullet_ab}">'
                         f'<w:multiLevelType w:val="hybridMultilevel"/>{bullet_lvls}</w:abstractNum>')
        ab_d = parse_xml(f'<w:abstractNum xmlns:w="{_W_NS}" w:abstractNumId="{decimal_ab}">'
                         f'<w:multiLevelType w:val="hybridMultilevel"/>{dec_lvls}</w:abstractNum>')
        # schema order: all abstractNum first, then num
        first_num = numbering.find(qn('w:num'))
        for ab in (ab_b, ab_d):
            if first_num is not None:
                first_num.addprevious(ab)
            else:
                numbering.append(ab)
        self._decimal_abstract_id = decimal_ab
        self._bullet_num_id = self._add_num(bullet_ab)

    def _add_num(self, abstract_id: int, restart: bool = False) -> int:
        from docx.oxml import parse_xml
        num_id = self._next_num_id
        self._next_num_id += 1
        override = ''
        if restart:
            override = ''.join(
                f'<w:lvlOverride w:ilvl="{i}"><w:startOverride w:val="1"/></w:lvlOverride>'
                for i in range(9))
        num = parse_xml(f'<w:num xmlns:w="{_W_NS}" w:numId="{num_id}">'
                        f'<w:abstractNumId w:val="{abstract_id}"/>{override}</w:num>')
        self._numbering.append(num)
        return num_id

    def _new_decimal_num(self) -> int:
        """Fresh w:num -> every numbered list starts again from 1."""
        return self._add_num(self._decimal_abstract_id, restart=True)

    @classmethod
    def _apply_heading_style(cls, p_elem, level: int) -> None:
        level = max(1, min(level, 6))
        pPr = p_elem.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            p_elem.insert(0, pPr)
        pStyle = pPr.find(qn('w:pStyle'))
        if pStyle is None:
            pStyle = OxmlElement('w:pStyle')
            cls._insert_in_order(pPr, pStyle, _PPR_ORDER)
        pStyle.set(qn('w:val'), f'Heading{level}')
        outline = pPr.find(qn('w:outlineLvl'))
        if outline is None:
            outline = OxmlElement('w:outlineLvl')
            cls._insert_in_order(pPr, outline, _PPR_ORDER)
        outline.set(qn('w:val'), str(level - 1))

    def _apply_base_alignment(self, p_elem) -> None:
        """FIX v7: Only apply base alignment if the paragraph doesn't already
        have an explicit w:jc element. This preserves original alignment from
        the template (e.g., RIGHT, inherited/None) instead of overwriting
        everything with base_alignment (JUSTIFY)."""
        a = self.base_alignment
        if a is None:
            return
        pPr = p_elem.find(qn('w:pPr'))
        # If paragraph already has explicit alignment, preserve it
        if pPr is not None:
            jc = pPr.find(qn('w:jc'))
            if jc is not None:
                return  # Keep existing alignment
        # Only set base alignment if no explicit alignment exists
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            p_elem.insert(0, pPr)
        val = a.xml_value if hasattr(a, 'xml_value') else str(a)
        jc = OxmlElement('w:jc')
        self._insert_in_order(pPr, jc, _PPR_ORDER)
        jc.set(qn('w:val'), val)

    # ------------------------------------------------------------------
    # process_document
    # ------------------------------------------------------------------

    @staticmethod
    def _distribute_lines(lines: List[str], word_counts: List[int]) -> List[int]:
        P = len(word_counts)
        if P == 0:
            return []
        L = len(lines)
        if L <= 0:
            return [0] * P
        total = sum(word_counts)
        if L >= P:
            # FIX v8: Give each paragraph 1 line, all extra lines to the
            # last paragraph. This ensures clones inherit the last
            # paragraph's alignment (e.g., jc=both) instead of a random
            # paragraph's alignment (e.g., jc=right from P[0]).
            alloc = [1] * P
            rem = L - sum(alloc)
            if rem > 0:
                alloc[-1] += rem
            return alloc
        if total <= 0:
            base, rem = divmod(L, P)
            return [base + (1 if i < rem else 0) for i in range(P)]
        raw = [L * w / total for w in word_counts]
        alloc = [int(x) for x in raw]
        rem = L - sum(alloc)
        order = sorted(range(P), key=lambda i: raw[i] - alloc[i], reverse=True)
        for i in order[:rem]:
            alloc[i] += 1
        return alloc


    # ------------------------------------------------------------------
    # Markdown table support
    # ------------------------------------------------------------------

    @staticmethod
    def _is_table_start(lines: List[str], idx: int) -> bool:
        """Check if lines starting at idx form a Markdown table (header + separator)."""
        if idx + 1 >= len(lines):
            return False
        header = lines[idx].strip()
        separator = lines[idx + 1].strip()
        if '|' not in header:
            return False
        if not _MD_TABLE_SEP_RE.match(separator):
            return False
        if '|' not in separator:
            return False
        # GFM: число колонок разделителя совпадает с заголовком
        return len(DocxWorker._parse_table_row(header)) == len(DocxWorker._parse_table_alignment(separator))

    @staticmethod
    def _parse_table_row(line: str) -> List[str]:
        """Parse a Markdown table row into cell values."""
        line = line.strip()
        if line.startswith('|'):
            line = line[1:]
        if line.endswith('|') and not line.endswith('\\|'):
            line = line[:-1]
        # \| внутри ячейки — экранированный символ, а не граница
        return [cell.strip().replace('\\|', '|') for cell in _MD_TABLE_CELL_SPLIT_RE.split(line)]

    @staticmethod
    def _parse_table_alignment(separator: str) -> List[str]:
        """Parse column alignments from a Markdown table separator row."""
        separator = separator.strip()
        if separator.startswith('|'):
            separator = separator[1:]
        if separator.endswith('|'):
            separator = separator[:-1]
        cols = separator.split('|')
        alignments = []
        for col in cols:
            col = col.strip()
            left = col.startswith(':')
            right = col.endswith(':')
            if left and right:
                alignments.append('center')
            elif right:
                alignments.append('right')
            else:
                alignments.append('left')
        return alignments

    @staticmethod
    def _split_text_by_tables(text: str) -> List[Dict[str, Any]]:
        """Split text into segments, some of which are Markdown tables."""
        lines = text.split('\n')
        segments: List[Dict[str, Any]] = []
        i = 0
        while i < len(lines):
            if DocxWorker._is_table_start(lines, i):
                # Collect table lines
                table_lines = [lines[i], lines[i + 1]]
                alignments = DocxWorker._parse_table_alignment(lines[i + 1])
                j = i + 2
                while j < len(lines) and '|' in lines[j].strip() and lines[j].strip():
                    table_lines.append(lines[j])
                    j += 1
                segments.append({
                    'type': 'table',
                    'lines': table_lines,
                    'header': DocxWorker._parse_table_row(table_lines[0]),
                    'rows': [DocxWorker._parse_table_row(l) for l in table_lines[2:]],
                    'alignments': alignments,
                })
                i = j
            else:
                text_lines: List[str] = []
                while i < len(lines):
                    if DocxWorker._is_table_start(lines, i):
                        break
                    text_lines.append(lines[i])
                    i += 1
                segments.append({
                    'type': 'text',
                    'lines': text_lines,
                })
        return segments

    def _create_table_element(self, header: List[str], rows: List[List[str]],
                              alignments: List[str], nsmap) -> Any:
        """Create a w:tbl element from parsed Markdown table data."""
        tbl = OxmlElement('w:tbl')

        # --- Table properties ---
        tblPr = OxmlElement('w:tblPr')
        tblW = OxmlElement('w:tblW')
        tblW.set(qn('w:w'), '5000')   # v13: full text width (was auto -> narrow table, wrapped words)
        tblW.set(qn('w:type'), 'pct')
        tblPr.append(tblW)

        # Borders
        tblBorders = OxmlElement('w:tblBorders')
        for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '4')
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), 'auto')
            tblBorders.append(border)
        tblPr.append(tblBorders)
        tbl.append(tblPr)

        # --- Table grid ---
        num_cols = max(len(header), max((len(r) for r in rows), default=0), 1)
        tblGrid = OxmlElement('w:tblGrid')
        col_width = str(_DEFAULT_TEXT_WIDTH_TWIPS // num_cols)
        for _ in range(num_cols):
            gridCol = OxmlElement('w:gridCol')
            # без ширины LibreOffice и часть версий Word схлопывают колонки
            gridCol.set(qn('w:w'), col_width)
            tblGrid.append(gridCol)
        tbl.append(tblGrid)

        def _make_cell(text: str, col_idx: int, bold: bool = False) -> Any:
            tc = OxmlElement('w:tc')
            tcPr = OxmlElement('w:tcPr')
            tc.append(tcPr)
            p = OxmlElement('w:p')
            pPr = OxmlElement('w:pPr')
            p.append(pPr)
            # Alignment
            if col_idx < len(alignments) and alignments[col_idx] != 'left':
                jc = OxmlElement('w:jc')
                jc.set(qn('w:val'), alignments[col_idx])
                self._insert_in_order(pPr, jc, _PPR_ORDER)
            self._append_inline(p, text, None, force_bold=bold)
            tc.append(p)
            return tc

        # --- Header row ---
        tr = OxmlElement('w:tr')
        trPr = OxmlElement('w:trPr')
        tblHeader = OxmlElement('w:tblHeader')
        tblHeader.set(qn('w:val'), 'true')
        trPr.append(tblHeader)
        tr.append(trPr)
        for col_idx in range(num_cols):
            cell_text = header[col_idx] if col_idx < len(header) else ''
            tr.append(_make_cell(cell_text, col_idx, bold=True))
        tbl.append(tr)

        # --- Data rows ---
        for row_data in rows:
            tr = OxmlElement('w:tr')
            for col_idx in range(num_cols):
                cell_text = row_data[col_idx] if col_idx < len(row_data) else ''
                tr.append(_make_cell(cell_text, col_idx, bold=False))
            tbl.append(tr)

        return tbl

    def _process_block(self, block: Dict[str, Any],
                       text_processor: Callable[[str, Dict[str, Any]], str],
                       nsmap) -> None:
        elements = [b['element'] for b in block['blocks']]
        if not elements:
            return
        word_counts = [b['word_count'] for b in block['blocks']]

        para_info = {
            'is_heading': False,
            'is_list_item': False,
            'style_name': block['blocks'][0]['para_style'].style_name,
            'alignment': block['blocks'][0]['para_style'].alignment,
            'word_count': block['word_count'],
            'paragraph_count': len(elements),
        }
        try:
            new_text = text_processor(block['text'], para_info)
        except Exception as e:
            # Ошибка на одном блоке не должна ронять весь документ
            logger.exception(f'DocxWorker: text_processor упал на блоке '
                             f'"{block["text"][:80]}": {e}')
            return
        if new_text is None:
            return
        if not isinstance(new_text, str):
            logger.warning(f'DocxWorker: text_processor вернул {type(new_text).__name__} '
                           f'вместо str — блок оставлен без изменений')
            return

        if new_text == block['text']:
            return  # nothing changed: keep original runs, numbering, alignment

        new_text = new_text.replace('\r\n', '\n').replace('\r', '\n')
        new_text = new_text.replace('\\n', '\n')
        # Пустые строки по краям ответа LLM превращались в пустые абзацы
        new_text = new_text.strip('\n')

        if new_text == block['text']:
            return

        self._prev_is_list = False
        self._cur_decimal = None

        # --- Check for Markdown tables in the new text ---
        segments = self._split_text_by_tables(new_text)
        has_tables = any(seg['type'] == 'table' for seg in segments)

        if not has_tables:
            # Original behaviour: no tables, distribute lines across paragraphs
            lines = new_text.split('\n')
            alloc = self._distribute_lines(lines, word_counts)

            pos = 0
            for elem, n in zip(elements, alloc):
                my_lines = lines[pos:pos + n]
                pos += n

                # шаблон для клонов нужен только если строк больше одной
                template = None
                if len(my_lines) > 1:
                    template = deepcopy(elem)
                    self._strip_unwanted_pPr_elements(template)

                self._apply_markdown_to_paragraph(elem, my_lines[0] if my_lines else '', nsmap)
                anchor = elem
                for line in my_lines[1:]:
                    new_p = deepcopy(template)
                    self._apply_markdown_to_paragraph(new_p, line, nsmap)
                    anchor.addnext(new_p)
                    anchor = new_p
            return

        # --- Table-aware path: process segments in order, inserting tables ---
        elem_idx = 0
        last_anchor = None
        first_template = deepcopy(elements[0])
        self._strip_unwanted_pPr_elements(first_template)

        for seg in segments:
            if seg['type'] == 'table':
                tbl = self._create_table_element(
                    seg['header'], seg['rows'], seg['alignments'], nsmap
                )
                if last_anchor is not None:
                    last_anchor.addnext(tbl)
                else:
                    elements[elem_idx].addprevious(tbl)
                last_anchor = tbl
            else:
                for line in seg['lines']:
                    if elem_idx < len(elements):
                        elem = elements[elem_idx]
                        elem_idx += 1
                        self._apply_markdown_to_paragraph(elem, line, nsmap)
                        last_anchor = elem
                    else:
                        new_p = deepcopy(first_template)
                        self._apply_markdown_to_paragraph(new_p, line, nsmap)
                        last_anchor.addnext(new_p)
                        last_anchor = new_p

        # Неиспользованные исходные абзацы очищаются, иначе в документе
        # остаётся старый текст рядом с новым (абзацы не удаляются: в них
        # может быть sectPr — граница раздела)
        for elem in elements[elem_idx:]:
            self._apply_markdown_to_paragraph(elem, '', nsmap)

    def process_document(self, input_path: str, output_path: str,
                         text_processor: Callable[[str, Dict[str, Any]], str],
                         skip_headings: bool = True,
                         skip_list_items: bool = False) -> None:
        self._validate_docx_path(input_path, 'input_path')
        if not output_path:
            raise ValueError('process_document: не задан output_path')
        if not callable(text_processor):
            raise TypeError('process_document: text_processor должен быть вызываемым')
        # Документ открывается напрямую из исходника — без промежуточной копии на диске
        doc = Document(input_path)
        # FIX v9: Ensure numbering definitions exist before processing
        self._ensure_numbering_definitions(doc)
        nsmap = _NSMAP
        raw_blocks = self._build_raw_blocks(doc, nsmap)
        blocks = self._merge_raw_blocks(raw_blocks, skip_headings, skip_list_items)
        for block in blocks:
            if block['is_table']:
                self._process_table(block['blocks'][0]['element'], text_processor,
                                    skip_headings, skip_list_items, nsmap)
                continue
            if skip_headings and block['type'] == 'heading':
                continue
            if skip_list_items and block['type'] == 'list_item':
                continue
            self._process_block(block, text_processor, nsmap)
        self._validate_document(doc)
        self._ensure_parent_dir(output_path)
        doc.save(output_path)

    def _validate_document(self, doc) -> None:
        body = doc.element.body
        nsmap = _NSMAP
        id_attr = qn('w:id')

        # --- Bookmarks: уникальные id и парность start/end ---
        starts = body.findall('.//w:bookmarkStart', namespaces=nsmap)
        ends_by_id: Dict[str, List[Any]] = {}
        for be in body.iterfind('.//w:bookmarkEnd', namespaces=nsmap):
            ends_by_id.setdefault(be.get(id_attr), []).append(be)
        max_id = max((i for i in (_to_int(bm.get(id_attr)) for bm in starts) if i is not None),
                     default=0)
        seen_ids = set()
        for bm in starts:
            bm_id = bm.get(id_attr)
            if bm_id not in seen_ids:
                seen_ids.add(bm_id)
                continue
            # дубликат: новый id и для start, и для «лишнего» end с тем же id
            max_id += 1
            new_id = str(max_id)
            bm.set(id_attr, new_id)
            seen_ids.add(new_id)
            same_ends = ends_by_id.get(bm_id, [])
            if len(same_ends) > 1:
                be = same_ends.pop()
                be.set(id_attr, new_id)
                ends_by_id.setdefault(new_id, []).append(be)
        start_ids = {bm.get(id_attr) for bm in starts}
        for bm in starts:
            if bm.get(id_attr) not in ends_by_id:
                bm.getparent().remove(bm)
        for be_id, ends in ends_by_id.items():
            if be_id not in start_ids:
                for be in ends:
                    be.getparent().remove(be)

        # --- Каждая ячейка таблицы обязана заканчиваться абзацем ---
        for tc in body.iterfind('.//w:tc', namespaces=nsmap):
            children = [c for c in tc if _local_tag(c) != 'tcPr']
            if not children or _local_tag(children[-1]) != 'p':
                tc.append(OxmlElement('w:p'))

        # --- sectPr тела документа должен быть последним элементом ---
        all_elements = list(body)
        if all_elements and _local_tag(all_elements[-1]) != 'sectPr':
            body_sect = body.find(qn('w:sectPr'))
            if body_sect is not None:
                body.append(body_sect)  # перемещение в конец
            else:
                # берём sectPr последнего раздела, а не первого
                for elem in reversed(all_elements):
                    if _local_tag(elem) != 'p':
                        continue
                    pPr = elem.find(qn('w:pPr'))
                    sectPr = pPr.find(qn('w:sectPr')) if pPr is not None else None
                    if sectPr is not None:
                        pPr.remove(sectPr)
                        body.append(sectPr)
                        break

    def _cell_raw_blocks(self, cell, nsmap) -> List[Dict[str, Any]]:
        raw = []
        for p in cell.findall('w:p', namespaces=nsmap):
            text = self._paragraph_text(p, nsmap)
            if self._is_heading_from_xml(p):
                btype = 'heading'
            elif self._is_list_item_from_xml(p):
                btype = 'list_item'
            else:
                btype = 'paragraph'
            raw.append({
                'text': text,
                'word_count': self._count_words(text),
                'para_style': self._extract_paragraph_style_from_xml(p),
                'runs': [],
                'type': btype,
                'element_index': -1,
                'element': p,
                'is_table': False,
                'has_placeholder': bool(_PLACEHOLDER_RE.search(text)),
            })
        return raw

    def _process_table(self, tbl_elem, text_processor,
                       skip_headings: bool, skip_list_items: bool, nsmap) -> None:
        # FIX v6: Process each cell independently to preserve cell boundaries.
        # Только прямые потомки: './/' захватывал строки и абзацы вложенных
        # таблиц, и они обрабатывались повторно. Вложенные таблицы — рекурсией.
        for row in tbl_elem.findall('w:tr', namespaces=nsmap):
            for cell in row.findall('w:tc', namespaces=nsmap):
                for nested in cell.findall('w:tbl', namespaces=nsmap):
                    self._process_table(nested, text_processor,
                                        skip_headings, skip_list_items, nsmap)
                raw = self._cell_raw_blocks(cell, nsmap)
                if not raw:
                    continue
                for block in self._merge_raw_blocks(raw, skip_headings, skip_list_items):
                    if skip_headings and block['type'] == 'heading':
                        continue
                    if skip_list_items and block['type'] == 'list_item':
                        continue
                    self._process_block(block, text_processor, nsmap)

    def find_context_document(self, input_path):
        self._validate_docx_path(input_path, 'input_path')
        doc = Document(input_path)
        nsmap = _NSMAP
        # Текст run'ов абзаца склеивается без разделителя, абзацы — через \n
        # (раньше было наоборот: слова рвались, а абзацы слипались)
        return '\n'.join(
            ''.join(t.text or '' for t in para.iterfind('.//w:t', namespaces=nsmap))
            for para in doc.element.body.iterfind('.//w:p', namespaces=nsmap)
        )

    @staticmethod
    def _snap_to_space(text: str, pos: int, window: int = 12) -> int:
        """Сдвигает точку разреза к ближайшему пробелу, чтобы не рвать слово
        между run'ами с разным форматированием."""
        if pos <= 0 or pos >= len(text) or text[pos - 1].isspace() or text[pos].isspace():
            return pos
        for delta in range(1, window + 1):
            for cand in (pos - delta, pos + delta):
                if 0 < cand < len(text) and text[cand - 1].isspace():
                    return cand
        return pos

    def _replace_text_in_t_elements(self, t_elements: List, new_text: str, original_lengths: List[int]) -> None:
        if not t_elements:
            return
        new_text = _sanitize_xml_text(new_text)
        if not new_text:
            for t in t_elements:
                t.text = ''
            return
        new_text_len = len(new_text)
        num_elements = len(t_elements)
        if len(original_lengths) != num_elements:
            raise ValueError('_replace_text_in_t_elements: длины original_lengths и t_elements не совпадают')
        total_original = sum(original_lengths)
        if total_original == 0 or num_elements == 1:
            t_elements[0].text = new_text
            if new_text != new_text.strip():
                t_elements[0].set(_XML_SPACE, 'preserve')
            for t in t_elements[1:]:
                t.text = ''
            return
        # Точки разреза пропорциональны исходным длинам, со сдвигом к пробелам
        cuts = []
        acc = 0
        for length in original_lengths[:-1]:
            acc += length
            cut = self._snap_to_space(new_text, round(new_text_len * acc / total_original))
            cuts.append(max(cut, cuts[-1] if cuts else 0))
        bounds = [0] + cuts + [new_text_len]
        for i, t in enumerate(t_elements):
            portion_text = new_text[bounds[i]:bounds[i + 1]]
            t.text = portion_text
            if portion_text and portion_text != portion_text.strip():
                t.set(_XML_SPACE, 'preserve')
