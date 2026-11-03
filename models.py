"""Модели данных и исключения (без зависимостей от Excel/Word/UI)."""
from dataclasses import dataclass


class SourceError(Exception):
    """Входной файл (students.xlsx / tickets.docx) отсутствует или непригоден."""


class JournalLockedError(Exception):
    """results.xlsx недоступен для записи (обычно открыт в Excel)."""


class JournalError(Exception):
    """results.xlsx повреждён или не читается по другой причине."""


class TicketNotFoundError(Exception):
    """Номер билета из журнала отсутствует в текущем tickets.docx."""


def norm(value) -> str:
    """Нормализация для сравнения: пробелы и регистр не важны."""
    return " ".join(("" if value is None else str(value)).split()).casefold()


@dataclass(frozen=True)
class Student:
    group: str
    last_name: str
    first_name: str

    @property
    def full_name(self) -> str:
        return f"{self.last_name} {self.first_name}"

    @property
    def key(self) -> tuple:
        return (norm(self.group), norm(self.last_name), norm(self.first_name))


@dataclass(frozen=True)
class Ticket:
    number: int
    questions: tuple


@dataclass(frozen=True)
class Assignment:
    ticket: Ticket
    repeat: bool
