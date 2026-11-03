"""Чтение входных данных: students.xlsx и tickets.docx."""
import logging
import re
from pathlib import Path
from typing import Iterable

from docx import Document
from openpyxl import load_workbook

from models import SourceError, Student, Ticket

log = logging.getLogger(__name__)

QUESTIONS_PER_TICKET = 3
HEADER_RE = re.compile(r"^Билет\s+(\d+)\s*$", re.IGNORECASE)
QUESTION_RE = re.compile(r"^(\d+)\s*[.)]\s*(.+)$")


def parse_tickets(paragraphs: Iterable) -> list:
    """Разбирает последовательность (текст, автонумерация) в список Ticket.

    Некорректные билеты пропускаются, причина пишется в лог.
    """
    tickets = {}
    number = None
    questions = []

    def finish():
        if number is None:
            return
        if len(questions) < QUESTIONS_PER_TICKET:
            log.error("Билет %s пропущен: вопросов %d из %d",
                      number, len(questions), QUESTIONS_PER_TICKET)
        elif number in tickets:
            log.error("Билет %s пропущен: номер уже встречался", number)
        else:
            if len(questions) > QUESTIONS_PER_TICKET:
                log.warning("Билет %s: лишние вопросы проигнорированы", number)
            tickets[number] = Ticket(number, tuple(questions[:QUESTIONS_PER_TICKET]))

    for text, numbered in paragraphs:
        text = text.strip()
        if not text:
            continue
        header = HEADER_RE.match(text)
        if header:
            finish()
            number, questions = int(header.group(1)), []
            continue
        if text.casefold().startswith("билет"):
            finish()
            number, questions = None, []
            log.error("Заголовок билета не распознан: %r — билет пропущен", text)
            continue
        if number is None:
            log.warning("Строка вне билета проигнорирована: %r", text)
            continue
        question = QUESTION_RE.match(text)
        if question:
            questions.append(question.group(2).strip())
        elif numbered:  # автоматическая нумерация Word: номера нет в тексте
            questions.append(text)
        else:
            log.warning("Билет %s: строка не похожа на вопрос: %r", number, text)
    finish()
    return [tickets[k] for k in sorted(tickets)]


def _is_numbered(paragraph) -> bool:
    ppr = paragraph._p.pPr
    if ppr is not None and ppr.numPr is not None:
        return True
    style = paragraph.style.name if paragraph.style is not None else ""
    return style.startswith("List")


def load_tickets(path) -> list:
    path = Path(path)
    if not path.exists():
        raise SourceError(f"Не найден файл со списком билетов: {path.name}")
    try:
        doc = Document(str(path))
    except Exception as exc:
        raise SourceError(f"Не удалось открыть {path.name}: {exc}") from exc
    tickets = parse_tickets((p.text, _is_numbered(p)) for p in doc.paragraphs)
    if not tickets:
        raise SourceError(f"В {path.name} нет ни одного корректного билета (см. lab2.log)")
    return tickets


def load_students(path) -> dict:
    """Возвращает {имя_группы: [Student, ...]}; имя листа = имя группы."""
    path = Path(path)
    if not path.exists():
        raise SourceError(f"Не найден файл со списком студентов: {path.name}")
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        raise SourceError(f"Не удалось открыть {path.name}: {exc}") from exc
    groups = {}
    try:
        for ws in wb.worksheets:
            students = []
            for row in ws.iter_rows(min_row=2, max_col=2, values_only=True):
                row = tuple(row) + (None,) * (2 - len(row))
                last, first = ("" if c is None else str(c).strip() for c in row)
                if not last and not first:
                    continue
                if not last or not first:
                    log.warning("Лист %r: неполная строка пропущена: %r", ws.title, row)
                    continue
                students.append(Student(ws.title, last, first))
            groups[ws.title] = students
    finally:
        wb.close()
    return groups
