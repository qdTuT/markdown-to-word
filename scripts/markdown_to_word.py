"""
Markdown到Word转换器
将Markdown文件转换为符合严格排版规范的Word（.docx）文档。
"""

import argparse
import glob
import re
import sys
import json
from copy import deepcopy
from pathlib import Path
from typing import Dict, List, Optional
try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Inches, Pt, RGBColor
except ImportError:
    raise ImportError(
        "markdown_to_word 需要 python-docx 库。请先安装：pip install python-docx"
    )

# ---------------------------------------------------------------------------
# 样式配置（默认规范）
# ---------------------------------------------------------------------------
STYLE_CONFIG = {
    'main_title': {
        'latin': '微软雅黑', 'eastasia': '微软雅黑',
        'size': 22, 'bold': True, 'color': '000000',
        'alignment': WD_ALIGN_PARAGRAPH.CENTER,
        'space_before': 12, 'space_after': 6,
        'line_spacing': 1.5
    },
    'chapter_title': {
        'latin': '黑体', 'eastasia': '黑体',
        'size': 16, 'bold': True, 'color': '000000',
        'alignment': WD_ALIGN_PARAGRAPH.LEFT,
        'space_before': 12, 'space_after': 2,
        'line_spacing': 1.5
    },
    'section_title': {
        'latin': '微软雅黑', 'eastasia': '微软雅黑',
        'size': 14, 'bold': True, 'color': '000000',
        'alignment': WD_ALIGN_PARAGRAPH.LEFT,
        'space_before': 8, 'space_after': 8,
        'line_spacing': 1.0
    },
    'subsection_title': {
        'latin': '黑体', 'eastasia': '黑体',
        'size': 12, 'bold': True, 'color': '000000',
        'alignment': WD_ALIGN_PARAGRAPH.LEFT,
        'space_before': 6, 'space_after': 3,
        'line_spacing': 1.0
    },
    'body': {
        'latin': 'Calibri', 'eastasia': '宋体',
        'size': 10.5, 'bold': False, 'color': '000000',
        'alignment': WD_ALIGN_PARAGRAPH.LEFT,
        'space_before': 0, 'space_after': 3,
        'line_spacing': 1.0
    },
    'list': {
        'latin': 'Calibri', 'eastasia': '宋体',
        'size': 10.5, 'bold': False, 'color': '000000',
        'alignment': WD_ALIGN_PARAGRAPH.LEFT,
        'space_before': 0, 'space_after': 2,
        'line_spacing': 1.0,
        'bullet_size': 'auto'
    },
    'code': {
        'latin': 'Consolas', 'eastasia': 'Consolas',
        'size': 10, 'bold': False, 'color': '000000',
        'background': 'F5F5F5',
        'inline_code': False
    },
    'blockquote': {
        'latin': 'Calibri', 'eastasia': '宋体',
        'size': 10.5, 'bold': False, 'italic': True,
        'color': '666666', 'background': 'E8E8E8',
        'border_color': '999999'
    },
    'table': {
        'latin': 'Calibri', 'eastasia': '宋体',
        'size': 10.5, 'bold': False, 'color': '000000'
    }
}

ALIGNMENT_MAP = {
    0: WD_ALIGN_PARAGRAPH.LEFT,
    1: WD_ALIGN_PARAGRAPH.CENTER,
    2: WD_ALIGN_PARAGRAPH.RIGHT
}

_CUSTOM_STYLES_CACHE: Optional[Dict] = None

def load_custom_styles() -> Dict:
    global _CUSTOM_STYLES_CACHE
    if _CUSTOM_STYLES_CACHE is not None:
        return _CUSTOM_STYLES_CACHE

    styles = STYLE_CONFIG.copy()
    config_file = Path(__file__).parent.parent / "custom_styles.json"
    
    if config_file.exists():
        try:
            with open(config_file, "r", encoding="utf-8") as f:
                custom_data = json.load(f)
                
            for key, val in custom_data.items():
                if key in styles and isinstance(val, dict):
                    # 处理 alignment 数值向 python-docx 枚举对象的转换
                    if "alignment" in val and isinstance(val["alignment"], int):
                        val["alignment"] = ALIGNMENT_MAP.get(val["alignment"], WD_ALIGN_PARAGRAPH.LEFT)
                    styles[key].update(val)
            print(f"✓ 已加载自定义配置文件：{config_file}")
        except Exception as e:
            print(f"警告：读取 custom_styles.json 失败（{e}），将使用默认样式")
            
    _CUSTOM_STYLES_CACHE = styles
    return styles

HEADING_MAP = {
    1: 'main_title',
    2: 'chapter_title',
    3: 'section_title',
    4: 'subsection_title',
    5: 'subsection_title',  # 更深层级统一视为子标题
    6: 'subsection_title',
}


# ---------------------------------------------------------------------------
# 文本清理辅助函数
# ---------------------------------------------------------------------------
# 中文字符 / 全角标点 / 半角符号（用于空格归一化）
_CJK_RANGE = r'\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff'
_CJK_PUNCT = '，。、；：！？…—～·“”‘’（）【】《》〈〉「」『』'
_ASCII_SYMBOLS = '!"#$%&\'()*+,-./:;<=>?@[\\]^_`{|}~'

# 参与「与中文之间去空格」的符号：下划线 `_` 除外。
# `_` 已不再作为粗体/斜体标记，属于普通字符，去掉两侧空格会把 `_文字_` 和中文粘连。
_SPACE_TRIM_SYMBOLS = _ASCII_SYMBOLS.replace('_', '')
_OP_CHARS = '=+-*/%^&|<>~$#'

# 单位/科学符号：中文语境下与数字、文字之间不留空格（如 `30 ℃` → `30℃`）
_UNIT_SYMBOLS = '℃℉％‰′″°±×÷≈≤≥≠∞→←↑↓↔⇒⇔⟶⟵'
# 这些符号后面即使跟英文字母也贴合（`A → B` → `A→B`、`25 ° C` → `25°C`）
_GLUE_LETTER_SYMBOLS = '°→←↑↓↔⇒⇔⟶⟵'
_SPACE = r'[ \t\u00a0\u3000]'
_ASCII_SYM_CLASS = re.escape(_SPACE_TRIM_SYMBOLS)
_CJK_PUNCT_CLASS = re.escape(_CJK_PUNCT)

# 文件路径 / URL / 邮箱：先占位，避免其中的括号被替换成中文括号
_PROTECT_RE = re.compile(
    r'[A-Za-z][A-Za-z0-9+.\-]*://[^\s）】《》“”]*'               # URL
    r'|[A-Za-z]:[\\/][^\s，。；：！？、（）【】《》“”]*'            # Windows 绝对路径
    r'|(?:\\\\|//|~/|\./|\.\./)[^\s，。；：！？、（）【】《》“”]*'  # UNC / 相对路径
    r'|[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}'        # 邮箱
)


def _has_cjk(text: str) -> bool:
    """判断文本中是否包含中文字符（用于决定是否做中文标点替换）。"""
    return re.search(f'[{_CJK_RANGE}]', text) is not None


def _is_cjk(ch: str) -> bool:
    """单个字符是否为中文（汉字）。"""
    return bool(ch) and re.match(f'[{_CJK_RANGE}]', ch) is not None


def _neighbor(text: str, index: int, step: int) -> str:
    """取 index 处字符前/后（跳过空白）的第一个字符，越界返回空串。

    跳过空白是为了兼容 `文生图 （Txt2Img）` 这类符号前后带空格的写法——
    这些空格随后会被 _normalize_spacing 去掉，不应影响「前后是否中文」的判断。
    """
    i = index + step
    while 0 <= i < len(text):
        ch = text[i]
        if ch not in ' \t\u00a0\u3000':
            return ch
        i += step
    return ''


def _mask_paths(text: str):
    """把路径/URL/邮箱替换为占位符，返回 (占位后文本, 原文列表)。"""
    store = []

    def _repl(match):
        store.append(match.group(0))
        return f'\x01{len(store) - 1}\x02'

    return _PROTECT_RE.sub(_repl, text), store


def _restore_paths(text: str, store) -> str:
    if not store:
        return text
    return re.sub(r'\x01(\d+)\x02', lambda m: store[int(m.group(1))], text)


def _convert_en_symbols(text: str, opening: bool = True):
    """按语境替换英文双引号与半角括号。

    替换条件（满足其一即替换，判断相邻字符时跳过空白）：
    1. 符号前面是中文；
    2. 符号后面是中文；
    3. 引号/括号**内部**含中文。
    三条都不满足时保持原样。

    成对处理：
    - 双引号成对判断：开引号之前、合引号之后、或引号内容含中文 → 整对换成 “ ”；
    - 半角括号成对判断：开括号之前、闭括号之后、或括号内容含中文 → 整对换成 （ ）；
    - 落单的引号/括号（配对不在当前文本片段内）只能按自身前后判断，避免出现半边替换。

    返回 (转换后文本, 引号开合状态)，便于同一段落被拆成多个 run 时继续配对。
    """
    if not text or ('"' not in text and '(' not in text and ')' not in text):
        return text, opening

    chars = list(text)

    def _should_convert(open_pos: int, close_pos: Optional[int] = None) -> bool:
        """前后有中文，或内容含中文，即需要替换。"""
        prev_ch = _neighbor(text, open_pos, -1)
        if _is_cjk(prev_ch):
            return True

        if close_pos is None:
            # 落单符号：没有配对，只能看另一侧相邻字符
            return _is_cjk(_neighbor(text, open_pos, 1))

        if _is_cjk(_neighbor(text, close_pos, 1)):
            return True
        return _has_cjk(text[open_pos + 1:close_pos])

    # ---- 双引号：先按出现顺序配对，再按语境决定整对是否替换 ----
    quotes = [i for i, ch in enumerate(text) if ch == '"']
    state = opening          # True 表示下一个引号是开引号
    pair_open = None         # 当前片段内已出现的开引号位置
    for pos in quotes:
        if state:
            pair_open = pos
            state = False
            continue

        # 当前引号作为合引号
        state = True
        if pair_open is not None:
            if _should_convert(pair_open, pos):
                chars[pair_open] = '“'
                chars[pos] = '”'
            pair_open = None
        else:
            # 开引号在上一个 run 中：只能按合引号自身的前后判断
            if _should_convert(pos):
                chars[pos] = '”'

    # 片段末尾仍未闭合的开引号：按自身前后判断
    if pair_open is not None and _should_convert(pair_open):
        chars[pair_open] = '“'

    # ---- 半角括号：成对匹配，按语境决定整对是否替换 ----
    stack = []
    for i, ch in enumerate(text):
        if ch == '(':
            stack.append(i)
        elif ch == ')' and stack:
            open_pos = stack.pop()
            if _should_convert(open_pos, i):
                chars[open_pos] = '（'
                chars[i] = '）'

    return ''.join(chars), state


def _normalize_spacing(text: str) -> str:
    """
    去除多余空格：
    1. 中文与英文/数字之间，以及中文与中文之间
    2. 中文与半角符号（含单双引号、反引号、路径分隔符等）之间
    3. 全角标点（含中文引号、括号）两侧
    4. 数字/字母与运算符之间（如 `10 %` → `10%`、`A / B` → `A/B`）
    5. 半角标点前后（如 `word ,` → `word,`）
    6. 单位/科学符号两侧（如 `30 ℃` → `30℃`），不区分中英文语境
    英文单词之间的空格、数字与数字之间的空格保持不变。
    """
    # 1. 中文 ↔ 英文/数字
    text = re.sub(rf'(?<=[{_CJK_RANGE}]){_SPACE}+(?=[A-Za-z0-9])', '', text)
    text = re.sub(rf'(?<=[A-Za-z0-9]){_SPACE}+(?=[{_CJK_RANGE}])', '', text)
    # 1b. 中文 ↔ 中文：汉字之间不留空格（`中文 中文` → `中文中文`）
    text = re.sub(rf'(?<=[{_CJK_RANGE}]){_SPACE}+(?=[{_CJK_RANGE}])', '', text)
    # 2. 中文 ↔ 半角符号
    text = re.sub(rf'(?<=[{_CJK_RANGE}]){_SPACE}+(?=[{_ASCII_SYM_CLASS}])', '', text)
    text = re.sub(rf'(?<=[{_ASCII_SYM_CLASS}]){_SPACE}+(?=[{_CJK_RANGE}])', '', text)
    # 3. 全角标点两侧
    text = re.sub(rf'{_SPACE}+(?=[{_CJK_PUNCT_CLASS}])', '', text)
    text = re.sub(rf'(?<=[{_CJK_PUNCT_CLASS}]){_SPACE}+', '', text)
    # 4. 数字/字母 ↔ 运算符
    text = re.sub(rf'(?<=[A-Za-z0-9]){_SPACE}+(?=[{re.escape(_OP_CHARS)}])', '', text)
    text = re.sub(rf'(?<=[{re.escape(_OP_CHARS)}]){_SPACE}+(?=[A-Za-z0-9])', '', text)
    # 5. 半角标点前后
    text = re.sub(rf'{_SPACE}+(?=[,.;:!?%\)\]\}}])', '', text)
    text = re.sub(rf'(?<=[\(\[\{{]){_SPACE}+', '', text)
    # 6. 单位/科学符号两侧（不区分语境）
    # 左侧：紧跟在数字或中文后时贴合（`30 ℃` → `30℃`、`90 %` → `90%`）
    text = re.sub(
        rf'(?<=[0-9{_CJK_RANGE}]){_SPACE}+(?=[{_UNIT_SYMBOLS}])', '', text)
    # 左侧：箭头/角度符号与英文字母也贴合（`A → B` → `A→B`）
    text = re.sub(
        rf'(?<=[A-Za-z]){_SPACE}+(?=[{_GLUE_LETTER_SYMBOLS}])', '', text)
    # 右侧：后接数字或中文时去除空格（`90 % 的` → `90%的`）；
    # 后接英文单词时保留空格，避免把 `30 ℃ and` 粘成 `30℃and`
    text = re.sub(
        rf'(?<=[{_UNIT_SYMBOLS}]){_SPACE}+(?=[0-9{_CJK_RANGE}])', '', text)
    # 右侧：箭头、角度符号后接字母同样贴合（`A → B` → `A→B`、`25 ° C` → `25°C`）
    text = re.sub(
        rf'(?<=[{_GLUE_LETTER_SYMBOLS}]){_SPACE}+(?=[A-Za-z0-9{_CJK_RANGE}])', '', text)
    return text


# ---------------------------------------------------------------------------
# 代码块注释清理
# ---------------------------------------------------------------------------
# 行注释标记（按语言）。只有出现在「行首」或「空白之后」的标记才视为注释起点，
# 这样 `a // b`（整除）、`x--`（自减）、`http://`（URL）不会被误判为注释。
_COMMENT_LINE_MARKERS = {
    'python': ('#',), 'py': ('#',), 'py3': ('#',), 'python3': ('#',),
    'bash': ('#',), 'sh': ('#',), 'shell': ('#',), 'zsh': ('#',),
    'ruby': ('#',), 'rb': ('#',), 'perl': ('#',), 'pl': ('#',),
    'r': ('#',), 'yaml': ('#',), 'yml': ('#',), 'toml': ('#',),
    'ini': ('#', ';'), 'conf': ('#',), 'dockerfile': ('#',),
    'powershell': ('#',), 'ps1': ('#',), 'makefile': ('#',), 'make': ('#',),
    'javascript': ('//',), 'js': ('//',), 'jsx': ('//',),
    'typescript': ('//',), 'ts': ('//',), 'tsx': ('//',),
    'java': ('//',), 'c': ('//',), 'cpp': ('//',), 'c++': ('//',),
    'csharp': ('//',), 'cs': ('//',), 'go': ('//',), 'golang': ('//',),
    'rust': ('//',), 'rs': ('//',), 'kotlin': ('//',), 'kt': ('//',),
    'swift': ('//',), 'scala': ('//',), 'php': ('//',), 'dart': ('//',),
    'groovy': ('//',), 'json5': ('//',), 'protobuf': ('//',), 'proto': ('//',),
    'sql': ('--',), 'lua': ('--',), 'haskell': ('--',), 'hs': ('--',),
    'vhdl': ('--',), 'elm': ('--',),
    'html': ('<!--',), 'xml': ('<!--',), 'markdown': ('<!--',), 'md': ('<!--',),
}

# 块注释定界符（按语言）。未列出的语言只处理行注释。
_COMMENT_BLOCK_PAIRS = {
    'c': (('/*', '*/'),), 'cpp': (('/*', '*/'),), 'c++': (('/*', '*/'),),
    'csharp': (('/*', '*/'),), 'cs': (('/*', '*/'),), 'java': (('/*', '*/'),),
    'javascript': (('/*', '*/'),), 'js': (('/*', '*/'),), 'jsx': (('/*', '*/'),),
    'typescript': (('/*', '*/'),), 'ts': (('/*', '*/'),), 'tsx': (('/*', '*/'),),
    'go': (('/*', '*/'),), 'golang': (('/*', '*/'),), 'rust': (('/*', '*/'),),
    'rs': (('/*', '*/'),), 'kotlin': (('/*', '*/'),), 'kt': (('/*', '*/'),),
    'swift': (('/*', '*/'),), 'scala': (('/*', '*/'),), 'php': (('/*', '*/'),),
    'dart': (('/*', '*/'),), 'css': (('/*', '*/'),), 'scss': (('/*', '*/'),),
    'less': (('/*', '*/'),), 'groovy': (('/*', '*/'),), 'protobuf': (('/*', '*/'),),
    'proto': (('/*', '*/'),), 'sql': (('/*', '*/'),), 'lua': (('--[[', ']]'),),
    'html': (('<!--', '-->'),), 'xml': (('<!--', '-->'),),
    'markdown': (('<!--', '-->'),), 'md': (('<!--', '-->'),),
}

# 未标注语言时的保守组合：只认行首标记，避免误伤 C 系预处理指令
_COMMENT_DEFAULT_LINE = ('#', '//', '--')
_COMMENT_DEFAULT_BLOCK = ()

# C 系预处理指令：行首 `#` 不是注释
_PREPROCESSOR_WORDS = (
    'include', 'define', 'undef', 'ifdef', 'ifndef', 'if', 'elif', 'else',
    'endif', 'pragma', 'error', 'warning', 'line', 'import',
)


def _comment_syntax(lang: Optional[str]):
    """按语言返回 (行注释标记, 块注释定界符对)。"""
    key = (lang or '').strip().lower()
    return (_COMMENT_LINE_MARKERS.get(key, _COMMENT_DEFAULT_LINE),
            _COMMENT_BLOCK_PAIRS.get(key, _COMMENT_DEFAULT_BLOCK))


def _clean_comment_text(text: str) -> str:
    """按正文规则清理注释正文，保留缩进与块注释行首 `*` 对齐符。"""
    out = []
    for line in text.split('\n'):
        if not line.strip():
            # 空行原样保留（多行块注释的换行缩进、闭合定界符的对齐依赖它）
            out.append(line)
            continue
        indent = line[:len(line) - len(line.lstrip())]
        rest = line.strip()
        prefix = ''
        if rest.startswith('*') and not rest.startswith('**'):
            prefix = '* '
            rest = rest[1:].lstrip()
        cleaned = clean_text(rest, is_code=False).strip()
        if not cleaned:
            out.append((indent + prefix).rstrip())
        else:
            out.append(indent + prefix + cleaned)
    return '\n'.join(out)


def _clean_line_comment(segment: str, marker: str) -> str:
    """清理单行注释：标记后统一保留一个空格。"""
    cleaned = _clean_comment_text(segment[len(marker):]).strip()
    return f'{marker} {cleaned}' if cleaned else marker


def _clean_block_comment(segment: str, opener: str, closer: str) -> str:
    """清理块注释内容，保留定界符与原有换行/缩进结构。"""
    inner = segment[len(opener):]
    if closer and inner.endswith(closer):
        inner = inner[:-len(closer)]
    cleaned = _clean_comment_text(inner)
    if not cleaned.strip():
        return opener + closer
    if '\n' in cleaned:
        return opener + cleaned + closer
    lead = ' ' if inner[:1] in (' ', '\t') else ''
    trail = ' ' if inner[-1:] in (' ', '\t') else ''
    return opener + lead + cleaned.strip() + trail + closer


def _match_line_marker(code: str, index: int, markers) -> Optional[str]:
    """判断 index 处是否为行注释起点（必须位于行首或空白之后）。"""
    prev = code[index - 1] if index else '\n'
    if prev not in ' \t\n':
        return None
    for marker in markers:
        if not code.startswith(marker, index):
            continue
        if marker == '#':
            rest = code[index + 1:].lstrip()
            # 跳过 shebang 与 C 系预处理指令
            if rest.startswith('!') or any(rest.startswith(w) for w in _PREPROCESSOR_WORDS):
                return None
        return marker
    return None


def _clean_code_comments(code: str, lang: Optional[str] = None) -> str:
    """清理代码块中的注释内容（代码本体保持原样）。

    仅对识别出的注释片段执行 clean_text；字符串字面量、URL、运算符等
    不会被误判。未标注语言时按「行首标记」保守处理。
    """
    line_markers, block_pairs = _comment_syntax(lang)
    openers = {open_tag: close_tag for open_tag, close_tag in block_pairs}
    block_openers = tuple(openers)          # 长定界符优先，避免 `--[[` 被 `--` 抢先
    quotes = ('"', "'", '`')

    out = []
    i, n = 0, len(code)
    seg_start = 0

    while i < n:
        ch = code[i]

        # 1. 字符串字面量整体跳过（含转义），其中的 # // 不当作注释
        if ch in quotes:
            j = i + 1
            while j < n:
                if code[j] == '\\':
                    j += 2
                    continue
                if code[j] == ch:
                    j += 1
                    break
                if code[j] == '\n':
                    break
                j += 1
            i = j
            continue

        # 2. 块注释
        opener = next((o for o in block_openers if code.startswith(o, i)), None)
        if opener:
            closer = openers[opener]
            end = code.find(closer, i + len(opener))
            out.append(code[seg_start:i])
            if end == -1:
                out.append(_clean_block_comment(code[i:], opener, ''))
                i = n
            else:
                out.append(_clean_block_comment(code[i:end + len(closer)], opener, closer))
                i = end + len(closer)
            seg_start = i
            continue

        # 3. 行注释
        marker = _match_line_marker(code, i, line_markers)
        if marker:
            end = code.find('\n', i)
            end = n if end == -1 else end
            out.append(code[seg_start:i])
            out.append(_clean_line_comment(code[i:end], marker))
            i = end
            seg_start = i
            continue

        i += 1

    out.append(code[seg_start:])
    return ''.join(out)


# LaTeX 常用箭头/符号 → Unicode（行内数学公式中常见）
LATEX_SYMBOL_MAP = {
    'longrightarrow': '⟶',
    'longleftarrow': '⟵',
    'rightarrow': '→',
    'leftarrow': '←',
    'leftrightarrow': '↔',
    'Rightarrow': '⇒',
    'Leftarrow': '⇐',
    'Leftrightarrow': '⇔',
    'uparrow': '↑',
    'downarrow': '↓',
    'to': '→',
    'gets': '←',
    'leq': '≤',
    'geq': '≥',
    'neq': '≠',
    'le': '≤',
    'ge': '≥',
    'ne': '≠',
    'times': '×',
    'div': '÷',
    'pm': '±',
    'approx': '≈',
    'infty': '∞',
}

# 按名称长度降序排列，避免短宏名抢先匹配长宏名（如 \le 与 \leftarrow）
_LATEX_MACRO_RE = re.compile(
    r'\\(' + '|'.join(
        re.escape(name) for name in sorted(LATEX_SYMBOL_MAP, key=len, reverse=True)
    ) + r')(?![a-zA-Z])'
)


def _replace_latex_symbols(text: str) -> str:
    """去除 LaTeX 数学定界符，并把常用宏替换为 Unicode 符号。"""
    # \( ... \) 、\[ ... \]，并兼容全角括号写法 \（ ... \）
    text = re.sub(r'\\\((.*?)\\\)', r'\1', text, flags=re.S)
    text = re.sub(r'\\\[(.*?)\\\]', r'\1', text, flags=re.S)
    text = re.sub(r'\\（(.*?)\\）', r'\1', text, flags=re.S)
    return _LATEX_MACRO_RE.sub(lambda m: LATEX_SYMBOL_MAP[m.group(1)], text)


def clean_text(text: str, is_code: bool = False,
                zh_context: Optional[bool] = None,
                quote_state: Optional[List[bool]] = None) -> str:
    """
    清理文本：
    - 去除所有 ` 符号（连同其前后的空格）
    - 非代码：将 LaTeX 常用符号（如 \\rightarrow）替换为 Unicode 字符
    - 非代码：把英文双引号、半角括号替换为中文引号、括号，
      但仅当符号前面或后面是中文时替换，前后都不是中文则保持原样
    - 非代码：去除数字、符号、引号、反引号、文件路径前后的空格，以及中文与英文/数字/标点之间的空格
    - 代码内容（is_code=True）：原样保留，不做标点替换和空格清理

    zh_context 为 None 时按文本自身是否含中文自动判断；
    行内多个 run（粗体/斜体等）应传入整行是否含中文，保证替换规则一致。
    quote_state 为长度 1 的列表，用于跨 run 记录引号开合状态（可选）。
    """
    if is_code:
        return text.replace('`', '')

    if zh_context is None:
        zh_context = _has_cjk(text)

    # LaTeX 行内公式符号 → Unicode（如 \rightarrow → →）
    text = _replace_latex_symbols(text)

    # 路径/URL/邮箱先占位，避免其中的半角括号被替换成中文括号
    text, protected = _mask_paths(text)

    if zh_context:
        opening = quote_state[0] if quote_state is not None else True
        text, opening = _convert_en_symbols(text, opening)
        if quote_state is not None:
            quote_state[0] = opening

    text = _restore_paths(text, protected)

    # 去除反引号及其前后空格，再做统一空格归一化
    text = re.sub(rf'{_SPACE}*`{_SPACE}*', '', text)
    text = _normalize_spacing(text)
    return text


def _clean_with_inline_code(text: str, zh_context: Optional[bool] = None) -> str:
    """整段清理时保护反引号内的代码（标题等只生成单个 run 的场景）。"""
    # 反引号前后的空格按需求去除，再分段处理
    text = re.sub(rf'{_SPACE}*(`[^`]+`){_SPACE}*', r'\1', text)
    parts = re.split(r'(`[^`]+`)', text)
    out = []
    for part in parts:
        if len(part) >= 2 and part.startswith('`') and part.endswith('`'):
            out.append(clean_text(part[1:-1], is_code=True))
        else:
            out.append(clean_text(part, is_code=False, zh_context=zh_context))
    return ''.join(out)


def _needs_join(left: str, right: str) -> bool:
    """判断相邻两个字符之间不应保留空格（用于跨 run 边界，规则与 _normalize_spacing 一致）。"""
    if not left or not right:
        return False
    left_cjk = re.match(f'[{_CJK_RANGE}]', left) is not None
    right_cjk = re.match(f'[{_CJK_RANGE}]', right) is not None
    if left_cjk and right_cjk:
        # 汉字之间一律贴合（与 _normalize_spacing 第 1b 条一致）
        return True
    if left_cjk or right_cjk:
        ascii_side = right if left_cjk else left
        if ascii_side.isalnum() or ascii_side in _SPACE_TRIM_SYMBOLS:
            return True
        return ascii_side in _CJK_PUNCT
    if left in _CJK_PUNCT or right in _CJK_PUNCT:
        return True
    if left in _UNIT_SYMBOLS:
        # 符号后接英文单词时保留空格（与 _normalize_spacing 第 6 条一致）
        return left in _GLUE_LETTER_SYMBOLS or not right.isalpha()
    if right in _UNIT_SYMBOLS:
        return True
    if left.isalnum() and right in _OP_CHARS:
        return True
    if left in _OP_CHARS and right.isalnum():
        return True
    if right in ',.;:!?%)]}':
        return True
    if left in '([{':
        return True
    return False


def _trim_run_spacing(paragraph) -> None:
    """中文与英文/标点分处不同 run（粗体、行内代码等）时，同样去掉交界处的空格。"""
    # 只处理普通空格：制表符是列表序号/复选框的结构分隔，不能删除
    edge_spaces = ' \u00a0\u3000'
    runs = [r for r in paragraph.runs if r.text]
    for left, right in zip(runs, runs[1:]):
        left_text, right_text = left.text, right.text
        if not left_text or not right_text:
            continue
        if left_text[-1] not in edge_spaces and right_text[0] not in edge_spaces:
            continue
        left_char = left_text.rstrip(edge_spaces)[-1:]
        right_char = right_text.lstrip(edge_spaces)[:1]
        if _needs_join(left_char, right_char):
            left.text = left_text.rstrip(edge_spaces)
            right.text = right_text.lstrip(edge_spaces)


# ---------------------------------------------------------------------------
# 底层格式化辅助函数
# ---------------------------------------------------------------------------
def set_run_font(run, latin_font: str, eastasia_font: str,
                    size: float, bold: bool = False, italic: bool = False,
                    color: Optional[str] = None) -> None:
    """设置 run 的字体（拉丁、东亚）及其它属性。"""
    run.font.name = latin_font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color)

    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn('w:rFonts'))
    if r_fonts is None:
        r_fonts = OxmlElement('w:rFonts')
        r_pr.append(r_fonts)
    r_fonts.set(qn('w:ascii'), latin_font)
    r_fonts.set(qn('w:hAnsi'), latin_font)
    r_fonts.set(qn('w:eastAsia'), eastasia_font)


# 中文双引号（U+201C / U+201D）
_QUOTE_CHARS = '“”'


def _set_run_latin_font(run_element, font: str) -> None:
    """把 run 的西文字体槽位（ascii / hAnsi / cs）都指向指定字体。

    U+201C、U+201D 属于「中英通用」字符，Word 默认用西文字体槽位（hAnsi）渲染，
    所以只设 eastAsia 没用，必须把 ascii/hAnsi 也改掉，并加 hint="eastAsia"。
    """
    r_pr = run_element.find(qn('w:rPr'))
    if r_pr is None:
        r_pr = OxmlElement('w:rPr')
        run_element.insert(0, r_pr)
    r_fonts = r_pr.find(qn('w:rFonts'))
    if r_fonts is None:
        r_fonts = OxmlElement('w:rFonts')
        r_pr.append(r_fonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs'):
        r_fonts.set(qn(attr), font)
    r_fonts.set(qn('w:hint'), 'eastAsia')


def _split_quote_runs(paragraph) -> None:
    """把中文双引号拆成独立 run，并改由中文字体渲染（不再使用 Calibri）。

    引号沿用所在 run 的东亚字体（正文为宋体、标题为黑体等），
    其余字符（英文、数字）仍保持原西文字体不变。
    """
    from docx.text.run import Run

    for run in list(paragraph.runs):
        text = run.text
        if not text or not any(ch in _QUOTE_CHARS for ch in text):
            continue

        segments = [seg for seg in re.split(f'([{_QUOTE_CHARS}])', text) if seg]
        if len(segments) < 2:
            continue

        # 取该 run 自己的东亚字体作为引号字体
        quote_font = '宋体'
        r_pr = run._element.find(qn('w:rPr'))
        if r_pr is not None:
            r_fonts = r_pr.find(qn('w:rFonts'))
            if r_fonts is not None:
                quote_font = r_fonts.get(qn('w:eastAsia')) or quote_font

        template = run._element
        anchor = template
        for seg in segments:
            new_el = deepcopy(template)
            for t in new_el.findall(qn('w:t')):
                new_el.remove(t)
            anchor.addnext(new_el)
            anchor = new_el

            new_run = Run(new_el, paragraph)
            new_run.text = seg
            # 只含引号的片段改用中文字体
            if all(ch in _QUOTE_CHARS for ch in seg):
                _set_run_latin_font(new_el, quote_font)

        paragraph._p.remove(template)


def apply_paragraph_style(paragraph, style_config: Dict) -> None:
    """应用段落的对齐、间距、行距。"""
    para = paragraph.paragraph_format
    para.alignment = style_config.get('alignment', WD_ALIGN_PARAGRAPH.LEFT)
    para.space_before = Pt(style_config.get('space_before', 0))
    para.space_after = Pt(style_config.get('space_after', 0))
    line_spacing = style_config.get('line_spacing')
    if line_spacing == 1.5:
        para.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    elif line_spacing == 2.0:
        para.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    else:
        para.line_spacing_rule = WD_LINE_SPACING.SINGLE


# ---------------------------------------------------------------------------
# 转换器
# ---------------------------------------------------------------------------
class MarkdownToWordConverter:
    def __init__(self, md_dir: Optional[Path] = None):
        self.md_dir = Path(md_dir) if md_dir else Path.cwd()

        # 检测 Pygments 是否可用
        self._pygments_available = False
        try:
            import pygments
            self._pygments_available = True
        except ImportError:
            pass

        self.styles = load_custom_styles()
        self.doc = Document()

        # 设置文档默认字体（虽然影响 Normal 样式，但可以实现手动覆盖所有段落）
        normal = self.doc.styles['Normal']
        normal.font.name = "Calibri"
        normal.font.size = Pt(10.5)
        r_fonts = normal.element.get_or_add_rPr().find(qn('w:rFonts'))
        if r_fonts is None:
            r_fonts = OxmlElement('w:rFonts')
            normal.element.get_or_add_rPr().append(r_fonts)
        r_fonts.set(qn('w:ascii'), 'Calibri')
        r_fonts.set(qn('w:hAnsi'), 'Calibri')
        r_fonts.set(qn('w:eastAsia'), '宋体')

    def _get_style(self, key: str) -> Dict:
        """获取样式配置，若不存在则返回 body 样式。"""
        return self.styles.get(key, self.styles['body'])

    def _last_block_is_heading(self) -> bool:
        """判断文档最后一个块级元素是否为标题段落（用于抑制标题后的空行）。"""
        body = self.doc.element.body
        last = None
        for child in reversed(list(body)):
            # 跳过节属性 sectPr，只取真正的块级内容
            if child.tag == qn('w:sectPr'):
                continue
            last = child
            break
        if last is None or last.tag != qn('w:p'):
            return False
        p_pr = last.find(qn('w:pPr'))
        if p_pr is None:
            return False
        p_style = p_pr.find(qn('w:pStyle'))
        if p_style is None:
            return False
        return (p_style.get(qn('w:val')) or '').startswith('Heading')

    # 无序列表多级符号配置
    BULLET_SYMBOLS = ['•', '◦', '▪', '▫']

    # ---- 主函数循环 ----
    def parse_markdown(self, content: str) -> None:
        lines = content.split('\n')
        i = 0

        list_active = False
        ordered_counters: Dict[int, int] = {}

        def reset_list_state():
            nonlocal list_active
            list_active = False
            ordered_counters.clear()

        def get_indent_level(line_str: str) -> int:
            # 提取前导空白，兼容 \t 与 \xa0 (不间断空格)
            leading = line_str[:len(line_str) - len(line_str.lstrip())]
            # 统一按「4 个空格 = 1 个 tab = 1 级缩进」归一化
            expanded = leading.replace('\t', '    ').replace('\xa0', ' ')
            width = len(expanded)
            if width == 0:
                return 0
            # 1~3 个空格也视为一级（兼容 2 空格缩进写法）；8 个空格 / 2 个 tab 为二级
            return max(1, width // 4)

        while i < len(lines):
            line = lines[i].rstrip()
            stripped = line.strip()

            if not stripped:
                # 标题后不允许紧跟空行：若上一个块级元素是标题，直接跳过该空段落
                if not self._last_block_is_heading():
                    para = self.doc.add_paragraph()
                    apply_paragraph_style(para, self._get_style('body'))
                reset_list_state()
                i += 1
                continue

            # 1. 标题
            if stripped.startswith('#'):
                self._add_header(line)
                reset_list_state()
                i += 1
                continue

            # 2. 水平分割线（直接去除整行，不渲染、不保留空行）
            if re.match(r'^[-*_]{3,}$', stripped):
                i += 1
                continue

            # 3. 图片
            if re.match(r'^!\[[^\]]*\]\([^)]+\)$', stripped):
                self._add_image(line)
                reset_list_state()
                i += 1
                continue

            # 4. 任务列表 (优先于普通无序列表判断)
            task_match = re.match(r'^[-*+]\s+\[([ xX])\]\s+(.*)$', stripped)
            if task_match:
                indent_level = get_indent_level(line)
                checked = task_match.group(1).lower() == 'x'
                item_text = task_match.group(2).strip()
                
                # 清除深层有序列表计数
                for lvl in list(ordered_counters.keys()):
                    if lvl >= indent_level:
                        del ordered_counters[lvl]

                self._add_list_item(item_text, indent_level=indent_level, 
                                    is_task=True, task_checked=checked)
                list_active = True
                i += 1
                continue

            # 5. 普通无序列表
            unord_match = re.match(r'^[-*+]\s+(.*)$', stripped)
            if unord_match:
                indent_level = get_indent_level(line)
                item_text = unord_match.group(1).strip()

                for lvl in list(ordered_counters.keys()):
                    if lvl >= indent_level:
                        del ordered_counters[lvl]

                self._add_list_item(item_text, ordered=False, indent_level=indent_level)
                list_active = True
                i += 1
                continue

            # 6. 有序列表（同时兼容 "1. " 与 "1) "）
            ord_match = re.match(r'^(\d+)[\.\)]\s+(.*)$', stripped)
            if ord_match:
                indent_level = get_indent_level(line)
                src_number = int(ord_match.group(1))
                item_text = ord_match.group(2).strip()

                for lvl in list(ordered_counters.keys()):
                    if lvl > indent_level:
                        del ordered_counters[lvl]

                # 尊重源文本显式编号：松散列表（条目间夹空白行/正文段落）会被
                # reset_list_state 清空计数，此时直接采用源编号即可正确渲染 1~6；
                # 仅在「紧邻的连续列表」且源编号仍为 1 时（如 1. 1. 1.）才自动递增。
                if indent_level in ordered_counters and src_number == 1:
                    counter = ordered_counters[indent_level] + 1
                else:
                    counter = src_number

                ordered_counters[indent_level] = counter
                self._add_list_item(item_text, ordered=True, indent_level=indent_level, counter=counter)
                list_active = True
                i += 1
                continue

            # 7. 代码块
            if stripped.startswith('```'):
                i = self._add_code_block(lines, i)
                reset_list_state()
                continue

            # 8. 引用块
            if stripped.startswith('>'):
                i = self._add_blockquote(lines, i)
                reset_list_state()
                continue

            # 9. 表格
            if '|' in line and i + 1 < len(lines) and '|' in lines[i + 1]:
                if re.match(r'^[\s\-\|:]+$', lines[i + 1].strip()):
                    i = self._add_table(lines, i)
                    reset_list_state()
                    continue

            # 10. 普通段落
            self._add_paragraph(line)
            reset_list_state()
            i += 1

        # 全文统一处理：中文双引号改用中文字体渲染（正文 Calibri 不再作用于引号）
        self._apply_quote_font()

    def _apply_quote_font(self) -> None:
        """遍历正文（含表格单元格）中的所有段落，把中文双引号切成中文字体 run。"""
        from docx.text.paragraph import Paragraph

        for p_el in list(self.doc.element.body.iter(qn('w:p'))):
            _split_quote_runs(Paragraph(p_el, self.doc))

    # ---- 标题 ----
    def _add_header(self, line: str) -> None:
        match = re.match(r'^(#{1,6})\s+(.+)$', line.strip())
        if not match:
            return
        level = len(match.group(1))
        text = match.group(2).strip()
        style_key = HEADING_MAP.get(level, 'subsection_title')
        style = self._get_style(style_key)

        heading_style = f'Heading {min(level, 6)}'
        para = self.doc.add_paragraph(style=heading_style)
        apply_paragraph_style(para, style)
        run = para.add_run()
        cleaned_text = _clean_with_inline_code(text)
        set_run_font(run,
                        latin_font=style['latin'],
                        eastasia_font=style['eastasia'],
                        size=style['size'],
                        bold=style.get('bold', True),
                        color=style.get('color'))
        run.text = cleaned_text

    # ---- 普通段落 ----
    def _add_paragraph(self, line: str) -> None:
        para = self.doc.add_paragraph()
        style = self._get_style('body')
        apply_paragraph_style(para, style)
        self._add_formatted_text(para, line, style_key='body')

    # ---- 带行内格式的文本 ----
    def _add_formatted_text(self, paragraph, text: str, style_key: str = 'body') -> None:
        """使用正则 Token 化高效解析 Markdown 行内格式（链接/粗体/斜体/代码）"""
        base_style = self._get_style(style_key)
        # 整行是否中文语境：决定引号/括号是否替换（避免英文段落被改成中文标点）
        zh_context = _has_cjk(text)
        # 跨 run 的引号开合状态，保证 "…" 被粗体/斜体拆开时仍能正确配对
        quote_state = [True]

        # 正则表达式匹配优先级：行内代码 > 链接 > 粗体 > 斜体
        # 仅支持星号写法：粗体 `**文字**`、斜体 `*文字*`；
        # 下划线写法 `__文字__` / `_文字_` 不再识别（保留原字符，避免误伤 file_name 这类标识符）
        pattern = re.compile(
            r'(`(?P<code>[^`]+)`) | '                             # 行内代码
            r'(\[(?P<link_text>[^\]]+)\]\((?P<link_url>[^)]+)\)) | '  # 链接 [text](url)
            r'(\*\*(?P<bold>[^*]+)\*\*) | '                       # 粗体 **文字**
            r'(\*(?P<italic>[^*]+)\*) ',                          # 斜体 *文字*
            re.VERBOSE
        )

        last_end = 0
        for match in pattern.finditer(text):
            start, end = match.span()

            # 处理匹配项之前的普通文本
            if start > last_end:
                plain_text = text[last_end:start]
                cleaned = clean_text(plain_text, is_code=False,
                                        zh_context=zh_context, quote_state=quote_state)
                run = paragraph.add_run(cleaned)
                set_run_font(run, base_style['latin'], base_style['eastasia'],
                                base_style['size'], bold=base_style.get('bold', False),
                                color=base_style.get('color'))

            # 分发匹配类型
            groups = match.groupdict()

            if groups['link_text'] is not None:
                self._add_hyperlink(paragraph, groups['link_url'], groups['link_text'], base_style)

            elif groups['code'] is not None:
                code_style = self._get_style('code')
                # 行内代码内容一律按「代码」处理：去掉首尾空格和反引号，
                # 但不做引号/括号替换（是否启用 inline_code 只影响样式，不影响内容）
                code_text = clean_text(groups['code'].strip(), is_code=True)
                if code_style.get('inline_code', True):
                    run = paragraph.add_run(code_text)
                    set_run_font(run, code_style.get('latin', 'Consolas'),
                                    code_style.get('eastasia', 'Consolas'),
                                    code_style.get('size', base_style['size']),
                                    color='000000')
                    self._run_shading(run, code_style.get('background', 'F5F5F5'))
                else:
                    # 未启用行内代码转换：按普通正文样式渲染，内容仍保持代码原样
                    run = paragraph.add_run(code_text)
                    set_run_font(run, base_style['latin'], base_style['eastasia'],
                                    base_style['size'], bold=base_style.get('bold', False),
                                    color=base_style.get('color'))

            elif groups['bold'] is not None:
                run = paragraph.add_run(
                    clean_text(groups['bold'], is_code=False,
                                zh_context=zh_context, quote_state=quote_state))
                set_run_font(run, base_style['latin'], base_style['eastasia'],
                             base_style['size'], bold=True, color=base_style.get('color'))

            elif groups['italic'] is not None:
                run = paragraph.add_run(
                    clean_text(groups['italic'], is_code=False,
                                zh_context=zh_context, quote_state=quote_state))
                set_run_font(run, base_style['latin'], base_style['eastasia'],
                                base_style['size'], italic=True, color=base_style.get('color'))

            last_end = end

        # 处理末尾剩余文本
        if last_end < len(text):
            cleaned = clean_text(text[last_end:], is_code=False,
                                    zh_context=zh_context, quote_state=quote_state)
            run = paragraph.add_run(cleaned)
            set_run_font(run, base_style['latin'], base_style['eastasia'],
                            base_style['size'], bold=base_style.get('bold', False),
                            color=base_style.get('color'))

        # 跨 run 的空格归一化（中文与英文/标点被粗体、行内代码等拆开时）
        _trim_run_spacing(paragraph)

    # ---- 列表项（支持有序、无序、任务列表） ----
    def _add_list_item(self, content: str, ordered: bool = False,
                        indent_level: int = 0, counter: Optional[int] = None,
                        is_task: bool = False, task_checked: bool = False) -> None:
        """设置列表段落与前缀渲染。"""
        style = self._get_style('list')
        para = self.doc.add_paragraph()
        apply_paragraph_style(para, style)

        # 无序列表符号字号：默认与列表正文字号一致（"auto"）。
        # 注意：单倍行距下 Word 的行高由「本行最大字号的 run」决定，若把符号放大到 14/16pt，
        # 整行行高会被抬高（Calibri 16pt 约 19.5pt，而 12pt 正文约 13~15pt），
        # 视觉上列表间距反而大于正文，与「列表 space_after 2pt < 正文 3pt」的配置相矛盾。
        # 如需放大符号，可把 bullet_size 改成数字（如 14）或按层级数组（如 [14, 16, 16]），
        # 但列表行高会随之变大。
        bullet_size_conf = style.get('bullet_size', 'auto')
        if isinstance(bullet_size_conf, (list, tuple)) and bullet_size_conf:
            bullet_size = bullet_size_conf[min(indent_level, len(bullet_size_conf) - 1)]
        elif isinstance(bullet_size_conf, (int, float)) and not isinstance(bullet_size_conf, bool):
            bullet_size = bullet_size_conf
        else:
            bullet_size = style['size']

        # 符号与文本之间的间距 (单位：英寸)
        # 无序列表符号后固定为一个空格，悬挂缩进按「符号 + 空格」宽度估算，保证换行后文本对齐；
        # 有序 / 任务列表沿用原有固定间距。
        if not is_task and not ordered:
            prefix_size = bullet_size
            hanging_gap = 0.6 * prefix_size / 72.0
        else:
            hanging_gap = 0.18

        para.paragraph_format.left_indent = Inches(0.25 * indent_level + hanging_gap)
        para.paragraph_format.first_line_indent = Inches(-hanging_gap)
        # =========================================================

        if is_task:
            self._add_checkbox(para, checked=task_checked)
            space_run = para.add_run('\t')  # 使用制表符 \t
            set_run_font(space_run, style['latin'], style['eastasia'],
                            style['size'], bold=False, color=style.get('color'))
        else:
            if ordered:
                prefix = f"{counter}.\t"
                # 保持正文默认字号
                prefix_size = style['size']
            else:
                symbol = self.BULLET_SYMBOLS[min(indent_level, len(self.BULLET_SYMBOLS) - 1)]
                # 无论源码中符号后有多少个空格，统一渲染为一个空格
                prefix = f"{symbol} "

            # 渲染前缀符号或序号
            run = para.add_run(prefix)
            set_run_font(run, style['latin'], style['eastasia'],
                            size=prefix_size, bold=False, color=style.get('color'))

        # 渲染正文文本（使用默认列表字号）
        self._add_formatted_text(para, content, style_key='list')

    # ---- 复选框 ----
    def _add_checkbox(self, paragraph, checked: bool) -> None:
        """添加可点击复选框（Word 内容控件 w14:checkbox）。

        关键点：w14:checkbox 的子元素与属性全部使用 **w14 命名空间**
        （`w14:checked` / `w14:checkedState` 的 `w14:val`、`w14:font`）。
        早期版本误写成 `w:val`，Word 读不到状态，点击时会退回内置的
        ☒/☐ 字形（表现为「点一下变叉、再点变空框」）。同时 sdtContent
        里的字形必须与 checkedState/uncheckedState 保持一致，否则
        首次显示与点击后显示会对不上。
        """
        from docx.oxml import parse_xml

        self._sdt_seq = getattr(self, '_sdt_seq', 0) + 1
        sdt_id = 1000 + self._sdt_seq
        check_val = '1' if checked else '0'
        shown_char = '\u2611' if checked else '\u2610'   # ☑ / ☐
        sdt_xml = (
            '<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
            '<w:sdtPr>'
            '<w:alias w:val="Checkbox"/>'
            '<w:tag w:val="checkbox"/>'
            f'<w:id w:val="{sdt_id}"/>'
            '<w14:checkbox>'
            f'<w14:checked w14:val="{check_val}"/>'
            '<w14:checkedState w14:val="2611" w14:font="Segoe UI Symbol"/>'
            '<w14:uncheckedState w14:val="2610" w14:font="Segoe UI Symbol"/>'
            '</w14:checkbox>'
            '</w:sdtPr>'
            '<w:sdtContent>'
            '<w:r><w:rPr>'
            '<w:rFonts w:ascii="Segoe UI Symbol" w:hAnsi="Segoe UI Symbol" '
            'w:eastAsia="Segoe UI Symbol" w:cs="Segoe UI Symbol"/>'
            '<w:sz w:val="21"/><w:szCs w:val="21"/>'
            '</w:rPr>'
            f'<w:t>{shown_char}</w:t></w:r>'
            '</w:sdtContent>'
            '</w:sdt>'
        )
        paragraph._p.append(parse_xml(sdt_xml))

    # ---- 代码块 ----
    def _add_code_block(self, lines: List[str], start_idx: int) -> int:
        i = start_idx + 1
        code_lines: List[str] = []
        first_line = lines[start_idx].strip()
        lang = None
        if first_line.startswith('```') and len(first_line) > 3:
            lang = first_line[3:].strip()
        while i < len(lines) and not lines[i].strip().startswith('```'):
            code_lines.append(lines[i])
            i += 1
        code = '\n'.join(code_lines)
        # 注释内容按正文规则清理空格与标点，代码本体保持原样
        code = _clean_code_comments(code, lang)
        para = self.doc.add_paragraph()
        para.paragraph_format.left_indent = Inches(0)
        para.paragraph_format.right_indent = Inches(0)
        para.paragraph_format.space_before = Pt(6)
        para.paragraph_format.space_after = Pt(6)
        code_style = self._get_style('code')
        bg = code_style.get('background')
        if bg:
            self._paragraph_shading(para, bg)

        highlight_success = False
        if lang and self._pygments_available:
            try:
                from pygments import highlight as pyg_highlight
                from pygments.lexers import get_lexer_by_name, guess_lexer
                from pygments.token import Token
                from pygments.styles import get_style_by_name

                try:
                    lexer = get_lexer_by_name(lang, stripall=False)
                except Exception:
                    lexer = guess_lexer(code)

                style = get_style_by_name('default')

                # 移除 Pygments 在结尾自动追加的换行符 \n
                tokens = list(lexer.get_tokens(code))
                while tokens and tokens[-1][1].rstrip('\r\n') == '':
                    tokens.pop()
                if tokens and tokens[-1][1].endswith('\n'):
                    t_type, t_val = tokens[-1]
                    tokens[-1] = (t_type, t_val.rstrip('\r\n'))

                for token_type, value in tokens:
                    if not value:
                        continue
                    run = para.add_run(value)
                    set_run_font(run,
                                    latin_font=code_style.get('latin', 'Consolas'),
                                    eastasia_font=code_style.get('eastasia', 'Consolas'),
                                    size=code_style.get('size', 10),
                                    color='000000')
                    self._apply_token_style(run, token_type, style)
                highlight_success = True
            except Exception as e:
                print(f"警告：语法高亮失败（{e}），使用纯文本")
                highlight_success = False

        if not highlight_success:
            run = para.add_run(code)
            set_run_font(run,
                            latin_font=code_style.get('latin', 'Consolas'),
                            eastasia_font=code_style.get('eastasia', 'Consolas'),
                            size=code_style.get('size', 10),
                            color=code_style.get('color', '000000'))

        return i + 1

    def _apply_token_style(self, run, token_type, style) -> None:
        try:
            token_style = style.style_for_token(token_type)
            if not token_style:
                return
            color = token_style.get('color')
            if color:
                run.font.color.rgb = RGBColor.from_string(color.lstrip('#'))
            if token_style.get('bold'):
                run.font.bold = True
            if token_style.get('italic'):
                run.font.italic = True
        except Exception:
            pass

    # ---- 引用块 ----
    def _add_blockquote(self, lines: List[str], start_idx: int) -> int:
        i = start_idx
        quote_lines: List[str] = []
        while i < len(lines) and lines[i].strip().startswith('>'):
            quote_lines.append(re.sub(r'^\s*>\s?', '', lines[i]))
            i += 1
        bq_s = self._get_style('blockquote')
        para = self.doc.add_paragraph()
        para.paragraph_format.left_indent = Inches(0)
        para.paragraph_format.right_indent = Inches(0)
        para.paragraph_format.space_before = Pt(6)
        para.paragraph_format.space_after = Pt(6)
        self._paragraph_shading(para, bq_s.get('background', 'E8E8E8'))
        self._paragraph_border(para, 'left', color=bq_s.get('border_color', '999999'))
        self._add_formatted_text(para, '\n'.join(quote_lines), style_key='blockquote')
        return i

    # ---- 表格 ----
    def _add_table(self, lines: List[str], start_idx: int) -> int:
        rows: List[List[str]] = []
        i = start_idx
        while i < len(lines) and '|' in lines[i]:
            if re.match(r'^[\s\-\|:]+$', lines[i].strip()):
                i += 1
                continue
            cells = [cell.strip() for cell in lines[i].strip().strip('|').split('|')]
            rows.append(cells)
            i += 1
        if not rows:
            return i
        n_cols = max(len(row) for row in rows)
        table = self.doc.add_table(rows=len(rows), cols=n_cols)
        table.style = 'Table Grid'
        for row_idx, row_data in enumerate(rows):
            for col_idx in range(n_cols):
                cell_data = row_data[col_idx] if col_idx < len(row_data) else ""
                cell = table.rows[row_idx].cells[col_idx]
                cell.text = ""
                para = cell.paragraphs[0]
                para.paragraph_format.space_before = Pt(3)
                para.paragraph_format.space_after = Pt(3)
                self._add_formatted_text(para, cell_data, style_key='table')
                if row_idx == 0:
                    for run in para.runs:
                        run.font.bold = True
        return i

    def _add_hyperlink(self, paragraph, url: str, text: str, base_style: Dict) -> None:
        """向段落中添加 Word 原生超链接（自动处理 URL 容错与空格保留）"""
        clean_url = url.strip()
        if not re.match(r'^(https?://|mailto:|ftp://|#|file://)', clean_url, re.IGNORECASE):
            clean_url = 'https://' + clean_url

        r_id = paragraph.part.relate_to(
            clean_url,
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
            is_external=True
        )

        hyperlink = OxmlElement('w:hyperlink')
        hyperlink.set(qn('r:id'), r_id)

        new_run = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')

        color = OxmlElement('w:color')
        color.set(qn('w:val'), '0000FF')
        rPr.append(color)

        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'single')
        rPr.append(u)

        rFonts = OxmlElement('w:rFonts')
        latin = base_style.get('latin', 'Calibri')
        eastasia = base_style.get('eastasia', '宋体')
        rFonts.set(qn('w:ascii'), latin)
        rFonts.set(qn('w:hAnsi'), latin)
        rFonts.set(qn('w:eastAsia'), eastasia)
        rPr.append(rFonts)

        sz_val = str(int(base_style.get('size', 10.5) * 2))
        for sz_tag in ('w:sz', 'w:szCs'):
            sz_el = OxmlElement(sz_tag)
            sz_el.set(qn('w:val'), sz_val)
            rPr.append(sz_el)

        new_run.append(rPr)

        text_node = OxmlElement('w:t')
        text_node.set(qn('xml:space'), 'preserve')
        text_node.text = clean_text(text, is_code=False)
        new_run.append(text_node)

        hyperlink.append(new_run)
        paragraph._p.append(hyperlink)

    # ---- 图片 ----
    def _add_image(self, line: str) -> None:
        match = re.match(r'^!\[([^\]]*)\]\(([^)]+)\)\s*$', line.strip())
        if not match:
            return
        alt_text = match.group(1).strip()
        img_path = Path(match.group(2).strip())
        if not img_path.is_absolute():
            img_path = self.md_dir / img_path
        if not img_path.exists():
            print(f"警告：找不到图片 {img_path}，跳过")
            return
        try:
            para = self.doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run()
            run.add_picture(str(img_path), width=Inches(6.0))
            if alt_text:
                caption = self.doc.add_paragraph(alt_text)
                cap_style = self._get_style('body')
                apply_paragraph_style(caption, cap_style)
                caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in caption.runs:
                    set_run_font(r,
                                    latin_font=cap_style['latin'],
                                    eastasia_font=cap_style['eastasia'],
                                    size=cap_style['size']-1,
                                    italic=True,
                                    color='666666')
        except Exception as e:
            print(f"警告：图片嵌入失败 {img_path}（{e}）")

    # ---- 辅助样式方法 ----
    def _run_shading(self, run, color: str) -> None:
        r_pr = run._element.get_or_add_rPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), color)
        r_pr.append(shd)

    def _paragraph_shading(self, paragraph, color: str) -> None:
        if not color:
            return
        p_pr = paragraph._p.get_or_add_pPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), color)
        p_pr.append(shd)

    def _paragraph_border(self, paragraph, edge: str = 'left',
                            color: str = '999999', size: str = '8') -> None:
        p_pr = paragraph._p.get_or_add_pPr()
        p_bdr = p_pr.find(qn('w:pBdr'))
        if p_bdr is None:
            p_bdr = OxmlElement('w:pBdr')
            p_pr.append(p_bdr)
        border = OxmlElement(f'w:{edge}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), size)
        border.set(qn('w:space'), '4')
        border.set(qn('w:color'), color)
        p_bdr.append(border)

    # ---- 目录 ----
    def add_table_of_contents(self, depth: int = 3) -> None:
        para = self.doc.add_paragraph()
        run = para.add_run()
        fld_begin = OxmlElement('w:fldChar')
        fld_begin.set(qn('w:fldCharType'), 'begin')
        instr = OxmlElement('w:instrText')
        instr.set(qn('xml:space'), 'preserve')
        instr.text = f' TOC \\o "1-{depth}" \\h \\z \\u '
        fld_sep = OxmlElement('w:fldChar')
        fld_sep.set(qn('w:fldCharType'), 'separate')
        placeholder = OxmlElement('w:t')
        placeholder.text = '（目录：请在Word中右键此处选择"更新域"生成）'
        fld_end = OxmlElement('w:fldChar')
        fld_end.set(qn('w:fldCharType'), 'end')
        run._r.append(fld_begin)
        run._r.append(instr)
        run._r.append(fld_sep)
        run._r.append(placeholder)
        run._r.append(fld_end)

    # ---- 保存 ----
    def save(self, output_path: str) -> None:
        self.doc.save(output_path)


# ---------------------------------------------------------------------------
# 顶层转换函数
# ---------------------------------------------------------------------------
def convert_markdown_to_word(input_file: str, output_file: Optional[str] = None,
                                include_toc: bool = False) -> str:
    input_path = Path(input_file)
    if not input_path.exists():
        raise FileNotFoundError(f"找不到输入文件：{input_file}")

    if output_file is None:
        output_file = str(input_path.with_suffix('.docx'))

    with open(input_path, 'r', encoding='utf-8') as f:
        md_content = f.read()

    converter = MarkdownToWordConverter(md_dir=input_path.parent)
    if include_toc:
        converter.add_table_of_contents()
    converter.parse_markdown(md_content)

    output_path = Path(output_file)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    converter.save(str(output_path))
    return str(output_path)


# ---------------------------------------------------------------------------
# 命令行入口
# ---------------------------------------------------------------------------
def main() -> int:
    # 检测路径是否因含空格被 Shell 拆分
    for i in range(len(sys.argv) - 2):
        curr, nxt = sys.argv[i+1], sys.argv[i+2]
        combined = curr + ' ' + nxt
        if (('/' in combined or '\\' in combined or ':' in combined) and
            ' ' in combined and not nxt.startswith('-')):
            if Path(combined).exists() or glob.glob(combined):
                print(f"错误：路径 '{combined}' 因包含空格被命令行拆分。")
                print("请在命令行中使用英文双引号包裹含空格的路径，例如：")
                print(f'  python {sys.argv[0]} "{combined}" ...')
                return 1

    parser = argparse.ArgumentParser(
        description='将Markdown文件转换为Word文档'
    )
    parser.add_argument('input_files', nargs='+', help='要转换的Markdown文件（支持通配符）')
    parser.add_argument('-o', '--output', help='输出Word文件路径（仅单文件转换时使用）')
    parser.add_argument('--output-dir', help='批量转换时的输出目录')
    parser.add_argument('--include-toc', action='store_true', help='在文档开头插入目录（打开文档后更新域生成）')

    args = parser.parse_args()

    # 支持位置参数形式的输出文件：python script.py input.md output.docx
    positional_output = None
    input_patterns = list(args.input_files)
    if len(input_patterns) == 2 and not args.output:
        first, second = input_patterns
        if not args.output_dir and second.lower().endswith('.docx'):
            glob_matches_first = glob.glob(first)
            glob_matches_second = glob.glob(second)
            if glob_matches_first and not glob_matches_second:
                positional_output = second
                input_patterns = [first]

    all_files: List[str] = []
    for pattern in input_patterns:
        all_files.extend(glob.glob(pattern))

    if not all_files:
        print("未找到匹配模式的文件")
        return 1

    output_dir = Path(args.output_dir) if args.output_dir else None
    if output_dir:
        output_dir.mkdir(parents=True, exist_ok=True)

    success_count = 0
    for input_file in all_files:
        try:
            input_path = Path(input_file)
            if len(all_files) == 1 and positional_output:
                output_file = positional_output
            elif len(all_files) == 1 and args.output:
                output_file = args.output
            elif output_dir:
                output_file = str(output_dir / (input_path.stem + ".docx"))
            else:
                output_file = None

            result = convert_markdown_to_word(
                input_file, output_file,
                include_toc=args.include_toc,
            )
            print(f"✓ 已转换：{input_file} → {result}")
            success_count += 1
        except Exception as e:
            print(f"✗ 转换失败 {input_file}：{e}")

    if success_count > 0:
        print(f"\n成功转换 {success_count}/{len(all_files)} 个文件")
    else:
        print("\n所有文件转换均失败")
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())