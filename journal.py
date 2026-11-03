"""Журнал результатов results.xlsx (только дописывание строк)."""
import os
from datetime import datetime
from pathlib import Path
from zipfile import BadZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from models import JournalError, JournalLockedError, Student

HEADER = ["Группа", "Фамилия", "Имя", "Номер билета", "Дата и время", "Повтор (да/нет)"]
DATE_FORMAT = "DD.MM.YYYY HH:MM:SS"


class ExcelJournal:
    def __init__(self, path):
        self.path = Path(path)

    def first_ticket_for(self, student: Student):
        """Номер билета из самой первой записи по студенту или None."""
        wb = self._load()
        if wb is None:
            return None
        for row in wb.active.iter_rows(min_row=2, max_col=4, values_only=True):
            if len(row) < 4:
                continue
            record = Student(*(("" if c is None else str(c)) for c in row[:3]))
            if record.key == student.key:
                try:
                    return int(row[3])
                except (TypeError, ValueError):
                    continue
        return None

    def append(self, student: Student, ticket_number: int, repeat: bool, when=None):
        wb = self._load()
        if wb is None:
            wb = Workbook()
            wb.active.append(HEADER)
            for col, width in zip("ABCDEF", (16, 18, 18, 14, 22, 16)):
                wb.active.column_dimensions[col].width = width
        ws = wb.active
        ws.append([
            student.group, student.last_name, student.first_name, ticket_number,
            when or datetime.now().replace(microsecond=0), "да" if repeat else "нет",
        ])
        ws.cell(row=ws.max_row, column=5).number_format = DATE_FORMAT
        self._save(wb)

    def _load(self):
        if not self.path.exists():
            return None
        try:
            return load_workbook(self.path)
        except PermissionError as exc:
            raise JournalLockedError(f"{self.path.name} занят другой программой") from exc
        except (BadZipFile, InvalidFileException, KeyError, OSError) as exc:
            raise JournalError(f"Не удалось прочитать {self.path.name}: {exc}") from exc

    def _save(self, wb):
        # Пишем во временный файл и подменяем оригинал: при сбое журнал остаётся целым.
        tmp = self.path.with_name(self.path.name + ".tmp")
        try:
            wb.save(tmp)
            os.replace(tmp, self.path)
        except PermissionError as exc:
            tmp.unlink(missing_ok=True)
            raise JournalLockedError(f"{self.path.name} занят другой программой") from exc
        except OSError as exc:
            tmp.unlink(missing_ok=True)
            raise JournalError(f"Не удалось записать {self.path.name}: {exc}") from exc
