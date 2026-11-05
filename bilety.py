#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Генератор экзаменационных билетов по шаблону Word.

Программа берёт:

* **шаблон билета** (Word-файл с таблицей и заполнителями вида ``<NUM>``,
  ``<DISCIP>``, ``<BILET_TASK>`` и т.п.);
* **базу вопросов** (Word-файл, где каждый вопрос — отдельный абзац);
* **файл конфигурации** (значения, которые нужно подставить вместо
  заполнителей, например ``<DISCIP> = ТАУ``);

и создаёт новый Word-файл с N билетами по M вопросов в каждом.

Вопросы в билет выбираются случайно (без повторов внутри билета).
Билеты верстаются так, чтобы на одной странице помещалось два билета,
а если два не влезают — один билет на страницу.

Все параметры задаются константами в блоке «НАСТРОЙКИ» ниже.
Запуск:  ``python bilety.py``
"""

from __future__ import annotations

import copy
import math
import random
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# =============================================================================
# НАСТРОЙКИ — правьте только этот блок
# =============================================================================

TEMPLATE_FILE = "Макет.docx"        # шаблон билета
QUESTIONS_FILE = "Вопросы.docx"     # база вопросов
CONFIG_FILE = "config.ini"          # значения для заполнителей
OUTPUT_FILE = "Билеты.docx"         # куда сохранить результат

N_TICKETS = 10                      # сколько билетов создать
QUESTIONS_PER_TICKET = 3            # сколько вопросов в каждом билете

MAX_TICKETS_PER_PAGE = 2            # 2 — два билета на страницу (если влезают), 1 — по одному
RANDOM_SEED: Optional[int] = None   # число — воспроизводимая случайная выборка, None — каждый раз новая
UNIQUE_TICKETS = True               # стараться не создавать билеты с одинаковым набором вопросов

QUESTION_NUMBERING = "{n}. {text}"  # как нумеровать вопросы внутри билета ("{text}" — без нумерации)
QUESTION_ALIGNMENT = "left"         # выравнивание вопросов: left / center / both
QUESTION_SPACE_AFTER_TWIPS = 60     # интервал после вопроса (1 см ≈ 567 twips)
QUESTION_HANGING_INDENT_TWIPS = 340 # висячий отступ для второй строки вопроса (0 — отключить)

# Технические допуски вёрстки (twips; 567 twips = 1 см)
PAGE_SLOT_PAD_TWIPS = 140           # запас по высоте на каждую «половину» страницы
LAST_ROW_SAFETY_TWIPS = 60          # запас при растягивании билета на половину страницы
HEIGHT_SAFETY_FACTOR = 1.05         # запас при оценке высоты «шапки» билета
VERBOSE = True                      # печатать отчёт о работе

#: Заполнители, которые подставляет сама программа (остальные берутся из файла конфигурации)
NUM_PLACEHOLDER = "<NUM>"
TASKS_PLACEHOLDER = "<BILET_TASK>"

# =============================================================================
# Служебные константы (менять не нужно)
# =============================================================================

PLACEHOLDER_RE = re.compile(r"<[A-ZА-ЯЁ][A-ZА-ЯЁ0-9_]*>")
HEADING_STYLE_HINTS = ("heading", "заголовок", "title", "название")
NUMBER_PREFIX_RE = re.compile(r"^\s*(?:\d{1,3}[.)]|[-–—•])\s+")
#: множитель «высота строки / размер шрифта» для популярных гарнитур
LINE_HEIGHT_FACTORS = {
    "calibri": 1.2207,
    "carlito": 1.2207,
    "times new roman": 1.1499,
    "liberation serif": 1.1499,
    "arial": 1.1499,
    "liberation sans": 1.1499,
    "cambria": 1.1700,
}
DEFAULT_LINE_HEIGHT_FACTOR = 1.20
DEFAULT_FONT_SIZE_PT = 12.0

#: приблизительная ширина символа в долях кегля (Calibri/Times, кириллица+латиница)
_NARROW_CHARS = "ijl.,;:'!|()[]{}«»/\\-–—"
_WIDE_CHARS = "mwMWШЩЖЮФМШфжющ"
DIGITS = "0123456789"
UPPER = "".join(chr(c) for c in range(ord("A"), ord("Z") + 1)) + "".join(
    chr(c) for c in range(ord("А"), ord("Я") + 1)
)


class BiletError(Exception):
    """Понятная пользователю ошибка (неверный шаблон, мало вопросов и т.п.)."""


# =============================================================================
# Данные
# =============================================================================


@dataclass
class StyleDefaults:
    """Значения по умолчанию из docDefaults шаблона (нужны для оценки высот)."""

    font_size_pt: float = DEFAULT_FONT_SIZE_PT
    space_before: float = 0.0
    space_after: float = 160.0
    line: float = 276.0
    line_rule: str = "auto"
    line_factor: float = DEFAULT_LINE_HEIGHT_FACTOR

    @property
    def single_line_twips(self) -> float:
        """Высота одной строки при одинарном интервале."""
        return self.font_size_pt * self.line_factor * 20.0


@dataclass
class Ticket:
    """Один билет: номер и тексты вопросов."""

    number: int
    questions: List[str]


# =============================================================================
# Чтение входных данных
# =============================================================================


def load_config(path: Path) -> Dict[str, str]:
    """Читает файл конфигурации с заполнителями.

    Поддерживаются строки вида ``<DISCIP> = ТАУ``, обычный INI
    (``DISCIP = ТАУ`` внутри секции ``[...]``), комментарии ``#`` и ``;``,
    а также значения, содержащие знак ``=``.
    """
    if not path.exists():
        raise BiletError(f"Файл конфигурации не найден: {path}")

    values: Dict[str, str] = {}
    for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.strip()
        if not line or line[0] in "#;" or (line.startswith("[") and line.endswith("]")):
            continue
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip().strip("[]").strip()
        if key.startswith("<") and key.endswith(">"):
            key = key[1:-1]
        value = value.strip().strip('"').strip("'")
        if key:
            values[key.upper()] = value
    return values


def load_questions(path: Path) -> List[str]:
    """Читает вопросы из Word-файла: каждый непустой абзац — один вопрос."""
    if not path.exists():
        raise BiletError(f"Файл с вопросами не найден: {path}")

    document = Document(str(path))
    questions: List[str] = []
    seen = set()

    for paragraph in document.paragraphs:
        style_name = (paragraph.style.name or "").lower()
        if any(hint in style_name for hint in HEADING_STYLE_HINTS):
            continue  # заголовки разделов пропускаем

        text = " ".join(paragraph.text.split())
        text = NUMBER_PREFIX_RE.sub("", text)  # убираем «1.» / «—», если номер вписан в текст
        if not text:
            continue

        key = text.casefold()
        if key in seen:  # одинаковые вопросы в базе не нужны
            continue
        seen.add(key)
        questions.append(text)

    if not questions:
        raise BiletError(
            f"В файле {path.name} не найдено ни одного вопроса. "
            "Каждый вопрос должен быть отдельным абзацем."
        )
    return questions


# =============================================================================
# Случайная выборка билетов
# =============================================================================


def build_tickets(
    questions: Sequence[str],
    n_tickets: int,
    per_ticket: int,
    rng: random.Random,
    unique: bool,
) -> List[Ticket]:
    """Формирует N билетов по M случайных вопросов (без повторов внутри билета)."""
    if n_tickets <= 0:
        raise BiletError("Количество билетов (N_TICKETS) должно быть больше нуля.")
    if per_ticket <= 0:
        raise BiletError("Количество вопросов в билете (QUESTIONS_PER_TICKET) должно быть больше нуля.")
    if per_ticket > len(questions):
        raise BiletError(
            f"В базе всего {len(questions)} вопрос(ов), а в билете требуется "
            f"{per_ticket}. Уменьшите QUESTIONS_PER_TICKET или дополните базу."
        )

    total_combinations = math.comb(len(questions), per_ticket)
    seen: set = set()
    tickets: List[Ticket] = []
    duplicated = 0

    for number in range(1, n_tickets + 1):
        indexes: Optional[Tuple[int, ...]] = None

        if unique and total_combinations > len(seen):
            for _ in range(300):
                candidate = tuple(sorted(rng.sample(range(len(questions)), per_ticket)))
                if candidate not in seen:
                    indexes = candidate
                    break
        if indexes is None:
            indexes = tuple(sorted(rng.sample(range(len(questions)), per_ticket)))
            if unique:
                duplicated += 1
        seen.add(indexes)
        tickets.append(Ticket(number=number, questions=[questions[i] for i in indexes]))

    if unique and duplicated:
        print(
            f"  ! Вариантов наборов меньше, чем билетов: {duplicated} билет(ов) "
            "совпадают по составу вопросов с другими.",
            file=sys.stderr,
        )
    return tickets


# =============================================================================
# Низкоуровневые помощники XML
# =============================================================================


#: порядок дочерних элементов w:pPr по схеме WordprocessingML
PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
    "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
    "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE",
    "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
    "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
    "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
    "sectPr", "pPrChange",
]
_PPR_INDEX = {name: index for index, name in enumerate(PPR_ORDER)}


def _insert_ppr_child(p_pr, element) -> None:
    """Вставляет элемент в w:pPr, соблюдая порядок, требуемый схемой OOXML."""
    tag = element.tag.split("}")[-1]
    position = _PPR_INDEX.get(tag, len(PPR_ORDER))
    for child in p_pr:
        child_tag = child.tag.split("}")[-1]
        if _PPR_INDEX.get(child_tag, len(PPR_ORDER)) > position:
            child.addprevious(element)
            return
    p_pr.append(element)


def _el(tag: str, **attrs: str):
    """Создаёт элемент WordprocessingML с атрибутами (ключи — имена без префикса)."""
    element = OxmlElement(tag)
    for name, value in attrs.items():
        element.set(qn(f"w:{name}"), str(value))
    return element


def _twips(value) -> int:
    """Переводит длину python-docx (EMU) в twips."""
    return int(round(value.twips))


def _normalize_values(values: Dict[str, str]) -> Dict[str, str]:
    """Приводит имена заполнителей к виду без угловых скобок и в верхний регистр."""
    normalized: Dict[str, str] = {}
    for key, value in values.items():
        normalized[key.strip().strip("<>").upper()] = str(value)
    return normalized


def _replace_placeholders(paragraph, values: Dict[str, str]) -> bool:
    """Заменяет заполнители в абзаце, сохраняя форматирование отдельных фрагментов.

    Заполнитель может быть «разорван» на несколько run'ов — в этом случае текст
    подставляется в первый run, а «хвосты» удаляются.
    """
    values = _normalize_values(values)
    runs = paragraph.findall(qn("w:r"))
    if not runs:
        return False

    texts: List[str] = []
    owners: List[int] = []
    for index, run in enumerate(runs):
        node = run.find(qn("w:t"))
        text = node.text if node is not None and node.text else ""
        texts.append(text)
        owners.extend([index] * len(text))

    full_text = "".join(texts)
    if "<" not in full_text:
        return False

    new_chars: List[Tuple[str, int]] = []
    position, changed = 0, False
    while position < len(full_text):
        match = PLACEHOLDER_RE.match(full_text, position)
        key = match.group(0)[1:-1].upper() if match else ""
        if match and key in values:
            for char in values[key]:
                new_chars.append((char, owners[position]))
            position = match.end()
            changed = True
        else:
            new_chars.append((full_text[position], owners[position]))
            position += 1
    if not changed:
        return False

    buckets: List[List[str]] = [[] for _ in runs]
    for char, run_index in new_chars:
        buckets[run_index].append(char)

    for index, run in enumerate(runs):
        new_text = "".join(buckets[index])
        node = run.find(qn("w:t"))
        if node is None:
            if not new_text:
                continue
            node = OxmlElement("w:t")
            run.append(node)
        node.text = new_text
        if new_text != new_text.strip():
            node.set(qn("xml:space"), "preserve")
    return True


def _paragraph_text(paragraph) -> str:
    return "".join(node.text or "" for node in paragraph.iter(qn("w:t")))


def _strip_volatile_ids(element) -> None:
    """Убирает w14:paraId/textId: при копировании они дублируются, Word это не любит."""
    for node in element.iter():
        for attr in ("{http://schemas.microsoft.com/office/word/2010/wordml}paraId",
                     "{http://schemas.microsoft.com/office/word/2010/wordml}textId"):
            if attr in node.attrib:
                del node.attrib[attr]


def _find_paragraph_with(element, placeholder: str):
    """Ищет абзац с заполнителем внутри элемента (например, внутри таблицы)."""
    for paragraph in element.iter(qn("w:p")):
        if placeholder in _paragraph_text(paragraph):
            return paragraph
    return None


def _ancestor_row(element):
    """Возвращает ближайший ``w:tr`` для элемента."""
    node = element
    while node is not None:
        if node.tag == qn("w:tr"):
            return node
        node = node.getparent()
    return None


def _set_row_height(row, twips: int) -> None:
    """Задаёт минимальную высоту строки (``atLeast``)."""
    tr_pr = row.find(qn("w:trPr"))
    if tr_pr is None:
        tr_pr = OxmlElement("w:trPr")
        row.insert(0, tr_pr)

    height = tr_pr.find(qn("w:trHeight"))
    if height is None:
        height = OxmlElement("w:trHeight")
        # в CT_TrPr trHeight идёт после cantSplit — вставляем в конец, порядок соблюдён
        tr_pr.append(height)
    height.set(qn("w:val"), str(int(twips)))
    height.set(qn("w:hRule"), "atLeast")


def _forbid_row_split(row) -> None:
    """Запрещает разрыв строки между страницами (``cantSplit``)."""
    tr_pr = row.find(qn("w:trPr"))
    if tr_pr is None:
        tr_pr = OxmlElement("w:trPr")
        row.insert(0, tr_pr)
    if tr_pr.find(qn("w:cantSplit")) is not None:
        return
    cant_split = OxmlElement("w:cantSplit")
    height = tr_pr.find(qn("w:trHeight"))
    if height is not None:
        height.addprevious(cant_split)
    else:
        tr_pr.append(cant_split)


def _set_cell_valign_top(cell) -> None:
    tc_pr = cell.find(qn("w:tcPr"))
    if tc_pr is None:
        tc_pr = OxmlElement("w:tcPr")
        cell.insert(0, tc_pr)
    valign = tc_pr.find(qn("w:vAlign"))
    if valign is None:
        valign = OxmlElement("w:vAlign")
        tc_pr.append(valign)
    valign.set(qn("w:val"), "top")


def _tiny_paragraph():
    """Почти невидимый пустой абзац (нужен после вложенной таблицы и между таблицами)."""
    paragraph = OxmlElement("w:p")
    p_pr = OxmlElement("w:pPr")
    p_pr.append(_el("w:spacing", before=0, after=0, line=2, lineRule="exact"))
    r_pr = OxmlElement("w:rPr")
    r_pr.append(_el("w:sz", val=2))
    r_pr.append(_el("w:szCs", val=2))
    p_pr.append(r_pr)
    paragraph.append(p_pr)
    return paragraph


# =============================================================================
# Оценка высоты — нужна, чтобы растянуть билет ровно на «половину страницы»
# =============================================================================


def read_style_defaults(document) -> StyleDefaults:
    """Читает docDefaults шаблона: кегль, интервалы, гарнитуру."""
    defaults = StyleDefaults()
    styles = document.styles.element
    doc_defaults = styles.find(qn("w:docDefaults"))
    if doc_defaults is None:
        return defaults

    r_pr = doc_defaults.find(qn("w:rPrDefault"))
    if r_pr is not None:
        r_pr = r_pr.find(qn("w:rPr"))
        if r_pr is not None:
            size = r_pr.find(qn("w:sz"))
            if size is not None:
                try:
                    defaults.font_size_pt = int(size.get(qn("w:val"))) / 2.0
                except (TypeError, ValueError):
                    pass

    p_pr = doc_defaults.find(qn("w:pPrDefault"))
    if p_pr is not None:
        p_pr = p_pr.find(qn("w:pPr"))
        spacing = p_pr.find(qn("w:spacing")) if p_pr is not None else None
        if spacing is not None:
            defaults.space_before = float(spacing.get(qn("w:before")) or 0)
            defaults.space_after = float(spacing.get(qn("w:after")) or 0)
            defaults.line = float(spacing.get(qn("w:line")) or 240)
            defaults.line_rule = spacing.get(qn("w:lineRule")) or "auto"

    defaults.line_factor = _detect_line_factor(document)
    return defaults


def _detect_line_factor(document) -> float:
    """Определяет множитель высоты строки по гарнитуре темы оформления."""
    try:
        for part in document.part.package.iter_parts():
            if str(part.partname).endswith("theme1.xml"):
                match = re.search(
                    r"<a:minorFont>.*?<a:latin typeface=\"([^\"]*)\"",
                    part.blob.decode("utf-8", "ignore"),
                    re.S,
                )
                if match:
                    name = match.group(1).strip().lower()
                    if name in LINE_HEIGHT_FACTORS:
                        return LINE_HEIGHT_FACTORS[name]
    except Exception:  # pragma: no cover - тема может отсутствовать
        pass
    return DEFAULT_LINE_HEIGHT_FACTOR


def _char_width_em(char: str) -> float:
    """Очень грубая оценка ширины символа в долях кегля."""
    if char == " ":
        return 0.226
    if char in DIGITS:
        return 0.51
    if char in _NARROW_CHARS:
        return 0.30
    if char in _WIDE_CHARS:
        return 0.85
    if char in UPPER:
        return 0.63
    return 0.50


def _count_lines(text: str, width_twips: float, font_size_pt: float) -> int:
    """Оценивает число строк, которое займёт текст в заданной ширине."""
    if not text:
        return 1
    em_twips = font_size_pt * 20.0
    # ширину слегка занижаем: пусть лучше оценка будет «с запасом вверх»
    available = max(width_twips * 0.97, em_twips)
    lines, current = 1, 0.0
    for word in text.split():
        word_width = sum(_char_width_em(ch) for ch in word) * em_twips
        space_width = _char_width_em(" ") * em_twips
        if current == 0.0:
            current = word_width
        elif current + space_width + word_width <= available:
            current += space_width + word_width
        else:
            lines += 1
            current = word_width
        # слишком длинное «слово» переносится принудительно
        while current > available:
            lines += 1
            current -= available
    return lines


def _paragraph_height(paragraph, defaults: StyleDefaults, width_twips: float) -> float:
    """Оценка высоты абзаца в twips (с учётом интервалов и переносов)."""
    p_pr = paragraph.find(qn("w:pPr"))
    before, after = defaults.space_before, defaults.space_after
    line, rule = defaults.line, defaults.line_rule
    font_size = defaults.font_size_pt

    if p_pr is not None:
        spacing = p_pr.find(qn("w:spacing"))
        if spacing is not None:
            before = float(spacing.get(qn("w:before")) or before)
            after = float(spacing.get(qn("w:after")) or after)
            line = float(spacing.get(qn("w:line")) or line)
            rule = spacing.get(qn("w:lineRule")) or rule

    sizes = [
        int(node.get(qn("w:val"))) / 2.0
        for node in paragraph.iter(qn("w:sz"))
        if node.get(qn("w:val"))
    ]
    if sizes:
        font_size = max(sizes)

    single = font_size * defaults.line_factor * 20.0
    if rule == "exact":
        line_height = line
    elif rule == "atLeast":
        line_height = max(line, single)
    else:  # auto — множитель к одинарному интервалу
        line_height = single * (line / 240.0)

    lines = _count_lines(_paragraph_text(paragraph), width_twips, font_size)
    return before + lines * line_height + after


def _grid_widths(table) -> List[int]:
    grid = table.find(qn("w:tblGrid"))
    if grid is None:
        return []
    widths = []
    for column in grid.findall(qn("w:gridCol")):
        try:
            widths.append(int(column.get(qn("w:w"))))
        except (TypeError, ValueError):
            widths.append(0)
    return widths


def _cell_gridspan(cell) -> int:
    tc_pr = cell.find(qn("w:tcPr"))
    if tc_pr is None:
        return 1
    span = tc_pr.find(qn("w:gridSpan"))
    if span is None:
        return 1
    try:
        return max(1, int(span.get(qn("w:val"))))
    except (TypeError, ValueError):
        return 1


def _vmerge_state(cell) -> Optional[str]:
    tc_pr = cell.find(qn("w:tcPr"))
    if tc_pr is None:
        return None
    merge = tc_pr.find(qn("w:vMerge"))
    if merge is None:
        return None
    return "restart" if merge.get(qn("w:val")) == "restart" else "continue"


def _cell_text_width(cell, column: int, grid_widths: Sequence[int], span: int) -> int:
    width = sum(grid_widths[column : column + span]) or 4000
    return max(600, width - 216)  # вычитаем поля ячейки по умолчанию (2 × 108)


def _cell_content_height(cell, defaults: StyleDefaults, text_width: float) -> float:
    return sum(_paragraph_height(p, defaults, text_width) for p in cell.findall(qn("w:p")))


def estimate_rows_height(rows: Sequence, defaults: StyleDefaults, grid_widths: Sequence[int]) -> List[float]:
    """Оценивает высоты строк таблицы с учётом объединённых по вертикали ячеек."""
    rows = list(rows)
    row_heights: List[float] = []
    spans: List[Tuple[int, int, float]] = []

    for index, row in enumerate(rows):
        tr_pr = row.find(qn("w:trPr"))
        minimum = 0.0
        if tr_pr is not None:
            height = tr_pr.find(qn("w:trHeight"))
            if height is not None:
                try:
                    minimum = float(height.get(qn("w:val")) or 0)
                except (TypeError, ValueError):
                    minimum = 0.0

        best_content = 0.0
        column = 0
        for cell in row.findall(qn("w:tc")):
            span = _cell_gridspan(cell)
            state = _vmerge_state(cell)
            text_width = _cell_text_width(cell, column, grid_widths, span)
            content = _cell_content_height(cell, defaults, text_width)
            if state == "restart":
                end = _find_span_end(rows, index, column)
                spans.append((index, end, content))
            elif state is None:
                best_content = max(best_content, content)
            column += span

        row_heights.append(max(minimum, best_content))

    for start, end, content in spans:
        total = sum(row_heights[start : end + 1])
        if content > total:
            row_heights[end] += content - total
    return row_heights


def _find_span_end(rows: Sequence, start: int, column: int) -> int:
    """Ищет последнюю строку вертикально объединённой ячейки, начинающейся в column."""
    end = start
    for index in range(start + 1, len(rows)):
        position = 0
        found_state = None
        for cell in rows[index].findall(qn("w:tc")):
            span = _cell_gridspan(cell)
            if position <= column < position + span:
                found_state = _vmerge_state(cell)
                break
            position += span
        if found_state != "continue":
            break
        end = index
    return end


# =============================================================================
# Сборка билета
# =============================================================================


def _question_paragraph(base_paragraph, text: str):
    """Создаёт абзац с вопросом на основе абзаца-заполнителя (формат сохраняется)."""
    paragraph = OxmlElement("w:p")

    p_pr = base_paragraph.find(qn("w:pPr"))
    p_pr = copy.deepcopy(p_pr) if p_pr is not None else OxmlElement("w:pPr")

    for tag in ("w:jc", "w:spacing", "w:ind"):
        node = p_pr.find(qn(tag))
        if node is not None:
            p_pr.remove(node)

    _insert_ppr_child(p_pr, _el("w:spacing", before=0, after=QUESTION_SPACE_AFTER_TWIPS))
    if QUESTION_HANGING_INDENT_TWIPS > 0:
        _insert_ppr_child(
            p_pr, _el("w:ind", left=QUESTION_HANGING_INDENT_TWIPS, hanging=QUESTION_HANGING_INDENT_TWIPS)
        )
    if QUESTION_ALIGNMENT in ("left", "center", "right", "both"):
        _insert_ppr_child(p_pr, _el("w:jc", val=QUESTION_ALIGNMENT))
    paragraph.append(p_pr)

    run = OxmlElement("w:r")
    base_run = base_paragraph.find(qn("w:r"))
    if base_run is not None:
        base_r_pr = base_run.find(qn("w:rPr"))
        if base_r_pr is not None:
            base_r_pr = copy.deepcopy(base_r_pr)
            for lang in base_r_pr.findall(qn("w:lang")):
                base_r_pr.remove(lang)  # чтобы проверка орфографии шла по-русски
            if len(base_r_pr):
                run.append(base_r_pr)
    text_node = OxmlElement("w:t")
    text_node.set(qn("xml:space"), "preserve")
    text_node.text = text
    run.append(text_node)
    paragraph.append(run)
    return paragraph


def _insert_questions(table, questions: Sequence[str], values: Dict[str, str]):
    """Заменяет абзац ``<BILET_TASK>`` на нумерованный список вопросов."""
    placeholder = _find_paragraph_with(table, TASKS_PLACEHOLDER)
    if placeholder is None:
        raise BiletError(
            f"В шаблоне не найден заполнитель {TASKS_PLACEHOLDER} — "
            "именно в это место подставляются вопросы."
        )

    for index, question in enumerate(questions, start=1):
        text = QUESTION_NUMBERING.format(n=index, text=question)
        text = " ".join(_replace_placeholders_text(text, values).split("\n"))
        placeholder.addprevious(_question_paragraph(placeholder, text))

    row = _ancestor_row(placeholder)
    placeholder.getparent().remove(placeholder)
    return row


def _replace_placeholders_text(text: str, values: Dict[str, str]) -> str:
    """Подставляет значения в строку (для случаев, когда заполнитель внутри вопроса)."""
    for key, value in _normalize_values(values).items():
        text = text.replace(f"<{key}>", value)
    return text


def _prepare_ticket_table(template_table, ticket: Ticket, values: Dict[str, str]):
    """Готовит копию таблицы-шаблона под конкретный билет."""
    table = copy.deepcopy(template_table)
    _strip_volatile_ids(table)

    # Вложенная таблица выравнивается по левому краю ячейки-контейнера
    tbl_pr = table.find(qn("w:tblPr"))
    if tbl_pr is not None:
        indent = tbl_pr.find(qn("w:tblInd"))
        if indent is not None:
            indent.set(qn("w:w"), "0")

    question_row = _insert_questions(table, ticket.questions, values)

    for paragraph in table.iter(qn("w:p")):
        _replace_placeholders(paragraph, values)

    for row in table.findall(qn("w:tr")):
        _forbid_row_split(row)
    return table, question_row


def _stretch_ticket(question_row, target_height: float, header_height: float) -> None:
    """Растягивает строку с вопросами, чтобы билет занял своё место на странице."""
    if question_row is None:
        return
    wanted = target_height - header_height - LAST_ROW_SAFETY_TWIPS - 20
    if wanted <= 0:
        return

    last_row = question_row
    tr_pr = last_row.find(qn("w:trPr"))
    current = 0.0
    if tr_pr is not None:
        height = tr_pr.find(qn("w:trHeight"))
        if height is not None:
            try:
                current = float(height.get(qn("w:val")) or 0)
            except (TypeError, ValueError):
                current = 0.0
    if wanted > current:
        _set_row_height(last_row, wanted)
    for cell in last_row.findall(qn("w:tc")):
        _set_cell_valign_top(cell)


def _container_table(total_width: int, indent: int):
    """Невидимая таблица-контейнер: одна колонка, одна строка на билет."""
    table = OxmlElement("w:tbl")

    tbl_pr = OxmlElement("w:tblPr")
    tbl_pr.append(_el("w:tblW", w=total_width, type="dxa"))
    tbl_pr.append(_el("w:jc", val="left"))
    tbl_pr.append(_el("w:tblInd", w=indent, type="dxa"))

    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(_el(f"w:{edge}", val="nil"))
    tbl_pr.append(borders)
    tbl_pr.append(_el("w:tblLayout", type="fixed"))

    margins = OxmlElement("w:tblCellMar")
    for edge in ("top", "left", "bottom", "right"):
        margins.append(_el(f"w:{edge}", w=0, type="dxa"))
    tbl_pr.append(margins)
    table.append(tbl_pr)

    grid = OxmlElement("w:tblGrid")
    grid.append(_el("w:gridCol", w=total_width))
    table.append(grid)
    return table


def _container_row(table, slot_height: float, width: int):
    """Строка контейнера: ровно «половина страницы» и без разрыва."""
    row = OxmlElement("w:tr")
    tr_pr = OxmlElement("w:trPr")
    tr_pr.append(OxmlElement("w:cantSplit"))
    height = OxmlElement("w:trHeight")
    height.set(qn("w:val"), str(int(slot_height)))
    height.set(qn("w:hRule"), "atLeast")
    tr_pr.append(height)
    row.append(tr_pr)

    cell = OxmlElement("w:tc")
    tc_pr = OxmlElement("w:tcPr")
    tc_pr.append(_el("w:tcW", w=width, type="dxa"))
    cell_margins = OxmlElement("w:tcMar")
    for edge in ("top", "left", "bottom", "right"):
        cell_margins.append(_el(f"w:{edge}", w=0, type="dxa"))
    tc_pr.append(cell_margins)
    tc_pr.append(_el("w:vAlign", val="top"))
    cell.append(tc_pr)
    row.append(cell)
    table.append(row)
    return cell


# =============================================================================
# Сборка документа
# =============================================================================


def build_document(
    template_path: Path,
    tickets: Sequence[Ticket],
    values: Dict[str, str],
) -> Tuple[Document, Dict[str, float]]:
    """Создаёт итоговый документ и возвращает его вместе с параметрами вёрстки."""
    if not template_path.exists():
        raise BiletError(f"Шаблон не найден: {template_path}")

    document = Document(str(template_path))
    if not document.tables:
        raise BiletError(
            f"В шаблоне {template_path.name} не найдена таблица билета. "
            "Билет должен быть оформлен таблицей Word."
        )

    template_table = document.tables[0]._tbl
    defaults = read_style_defaults(document)
    section = document.sections[0]

    # --- сколько места есть на странице и сколько отдаём одному билету ---
    usable_height = (
        _twips(section.page_height) - _twips(section.top_margin) - _twips(section.bottom_margin)
    )
    per_page = max(1, int(MAX_TICKETS_PER_PAGE))
    slot_height = usable_height / per_page - PAGE_SLOT_PAD_TWIPS
    if slot_height <= 500:
        raise BiletError("Слишком маленькая страница в шаблоне — билет не помещается.")

    grid_widths = _grid_widths(template_table)

    # --- геометрия контейнера: копируем ширину и отступ шаблонной таблицы ---
    tbl_pr = template_table.find(qn("w:tblPr"))
    total_width, indent = 9000, 0
    if tbl_pr is not None:
        width_node = tbl_pr.find(qn("w:tblW"))
        if width_node is not None and width_node.get(qn("w:w")):
            total_width = int(width_node.get(qn("w:w")))
        indent_node = tbl_pr.find(qn("w:tblInd"))
        if indent_node is not None and indent_node.get(qn("w:w")):
            indent = int(indent_node.get(qn("w:w")))

    # --- очищаем тело документа, оставляя свойства раздела ---
    body = document.element.body
    sect_pr = body.find(qn("w:sectPr"))
    for child in list(body):
        if child is not sect_pr:
            body.remove(child)

    container = _container_table(total_width, indent)
    unknown_placeholders: set = set()
    header_height: Optional[float] = None
    natural_heights: List[float] = []

    for ticket in tickets:
        cell = _container_row(container, slot_height, total_width)

        # <NUM> у каждого билета свой, остальные значения — из файла конфигурации
        ticket_values = dict(values)
        ticket_values["NUM"] = str(ticket.number)
        table, question_row = _prepare_ticket_table(template_table, ticket, ticket_values)

        unknown_placeholders |= {
            match.group(0)
            for paragraph in table.iter(qn("w:p"))
            for match in PLACEHOLDER_RE.finditer(_paragraph_text(paragraph))
        }

        rows = table.findall(qn("w:tr"))
        if header_height is None:
            # высота «шапки» = строки над строкой с вопросами; считаем один раз
            # и с запасом вверх — тогда растянутый билет гарантированно не выше своей доли
            header_index = rows.index(question_row) if question_row in rows else max(len(rows) - 1, 0)
            header_rows = rows[:header_index]
            header_height = (
                sum(estimate_rows_height(header_rows, defaults, grid_widths)) * HEIGHT_SAFETY_FACTOR
                + 30 * len(header_rows)
            )

        natural_heights.append(sum(estimate_rows_height(rows, defaults, grid_widths)))
        _stretch_ticket(question_row, slot_height, header_height)
        cell.append(table)
        cell.append(_tiny_paragraph())  # ячейка обязана заканчиваться абзацем

    if unknown_placeholders:
        print(
            "  ! Остались незаполненные заполнители (нет значения в "
            f"{CONFIG_FILE}): " + ", ".join(sorted(unknown_placeholders)),
            file=sys.stderr,
        )

    if sect_pr is not None:
        sect_pr.addprevious(container)
        sect_pr.addprevious(_tiny_paragraph())
    else:  # pragma: no cover - раздел есть почти всегда
        body.append(container)

    layout = {
        "usable_height": float(usable_height),
        "slot_height": float(slot_height),
        "header_height": float(header_height or 0),
        "per_page": float(per_page),
        "max_natural_height": float(max(natural_heights) if natural_heights else 0),
    }
    return document, layout


# =============================================================================
# Отчёт
# =============================================================================


def print_report(
    tickets: Sequence[Ticket],
    questions: Sequence[str],
    layout: Dict[str, float],
    output_path: Path,
) -> None:
    """Печатает краткий отчёт о том, что получилось."""
    used = {question for ticket in tickets for question in ticket.questions}
    per_page = int(layout["per_page"])
    usable = layout["usable_height"]
    slot = layout["slot_height"]
    heading = layout["header_height"]
    natural = layout["max_natural_height"]

    print("Готово.")
    print(f"  Файл:                  {output_path}")
    print(f"  Билетов:               {len(tickets)}")
    print(f"  Вопросов в билете:     {len(tickets[0].questions) if tickets else 0}")
    print(f"  Вопросов в базе:       {len(questions)} (использовано {len(used)})")
    print(f"  Билетов на странице:   до {per_page} (если не влезают — 1)")
    print(f"  Высота полосы набора:  {usable:.0f} twips")
    print(f"  Место под 1 билет:     {slot:.0f} twips (шапка ≈ {heading:.0f})")
    print(f"  Оценка высоты билета:  {natural:.0f} twips")
    if per_page > 1 and natural > slot:
        print(
            "  ! Билет выглядит выше отведённого места: на страницах будет по одному билету. "
            "Уменьшите QUESTIONS_PER_TICKET или QUESTION_SPACE_AFTER_TWIPS "
            "(или поставьте MAX_TICKETS_PER_PAGE = 1)."
        )


# =============================================================================
# Точка входа
# =============================================================================


def _configure_console() -> None:
    """Чтобы русский текст в консоли Windows не ломал программу."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError):  # pragma: no cover
            pass


def main() -> int:
    _configure_console()
    base = Path(__file__).resolve().parent
    template_path = base / TEMPLATE_FILE
    questions_path = base / QUESTIONS_FILE
    config_path = base / CONFIG_FILE
    output_path = base / OUTPUT_FILE

    try:
        config = load_config(config_path)
        questions = load_questions(questions_path)
        rng = random.Random(RANDOM_SEED)
        tickets = build_tickets(questions, N_TICKETS, QUESTIONS_PER_TICKET, rng, UNIQUE_TICKETS)

        values = dict(config)
        document, layout = build_document(template_path, tickets, values)
        document.save(str(output_path))
    except BiletError as error:
        print(f"Ошибка: {error}", file=sys.stderr)
        return 1

    if VERBOSE:
        print_report(tickets, questions, layout, output_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
