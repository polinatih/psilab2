import logging
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest import mock

from docx import Document
from openpyxl import Workbook, load_workbook

import journal as journal_module
from journal import ExcelJournal
from models import JournalLockedError, SourceError, Student, TicketNotFoundError
from service import TicketService
from sources import load_students, load_tickets, parse_tickets

logging.disable(logging.CRITICAL)


def lines(*texts):
    return [(t, False) for t in texts]


def ticket_lines(n, count=3):
    return lines(f"Билет {n}", *[f"{i}. Вопрос {i}" for i in range(1, count + 1)])


class ParseTicketsTest(unittest.TestCase):
    def test_valid_tickets(self):
        tickets = parse_tickets(ticket_lines(1) + lines("") + ticket_lines(2))
        self.assertEqual([t.number for t in tickets], [1, 2])
        self.assertEqual(tickets[0].questions, ("Вопрос 1", "Вопрос 2", "Вопрос 3"))

    def test_ticket_with_less_than_3_questions_is_skipped(self):
        tickets = parse_tickets(ticket_lines(1) + ticket_lines(2, count=2) + ticket_lines(3))
        self.assertEqual([t.number for t in tickets], [1, 3])

    def test_unrecognized_header_is_skipped_with_its_questions(self):
        tickets = parse_tickets(ticket_lines(1) + lines("Билет X", "1. a", "2. b", "3. c")
                                + ticket_lines(2))
        self.assertEqual([t.number for t in tickets], [1, 2])
        self.assertEqual(tickets[0].questions, ("Вопрос 1", "Вопрос 2", "Вопрос 3"))

    def test_duplicate_number_keeps_first(self):
        tickets = parse_tickets(ticket_lines(1) + lines("Билет 1", "1. x", "2. y", "3. z"))
        self.assertEqual(len(tickets), 1)
        self.assertEqual(tickets[0].questions[0], "Вопрос 1")

    def test_load_from_real_docx(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "tickets.docx"
            doc = Document()
            for text, _ in ticket_lines(1) + ticket_lines(2, count=1) + ticket_lines(3):
                doc.add_paragraph(text)
            doc.add_paragraph("Билет 4")
            for q in ("раз", "два", "три"):  # автонумерация Word
                doc.add_paragraph(q, style="List Number")
            doc.save(path)
            tickets = load_tickets(path)
        self.assertEqual([t.number for t in tickets], [1, 3, 4])
        self.assertEqual(tickets[2].questions, ("раз", "два", "три"))

    def test_missing_files_raise_source_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(SourceError):
                load_tickets(Path(tmp) / "nope.docx")
            with self.assertRaises(SourceError):
                load_students(Path(tmp) / "nope.xlsx")


class StudentsTest(unittest.TestCase):
    def test_groups_and_empty_group(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "students.xlsx"
            wb = Workbook()
            wb.active.title = "A-1"
            wb.active.append(["Фамилия", "Имя"])
            wb.active.append(["Иванов", " Иван "])
            wb.active.append([None, None])
            wb.active.append(["Без имени", None])
            wb.create_sheet("B-2").append(["Фамилия", "Имя"])
            wb.save(path)
            groups = load_students(path)
        self.assertEqual(list(groups), ["A-1", "B-2"])
        self.assertEqual([s.full_name for s in groups["A-1"]], ["Иванов Иван"])
        self.assertEqual(groups["B-2"], [])


class ServiceTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name) / "results.xlsx"
        self.journal = ExcelJournal(self.path)
        self.tickets = parse_tickets(sum((ticket_lines(n) for n in range(1, 6)), []))
        self.service = TicketService(self.journal, self.tickets)
        self.ivan = Student("A-1", "Иванов", "Иван")

    def tearDown(self):
        self.tmp.cleanup()

    def rows(self):
        return list(load_workbook(self.path).active.iter_rows(values_only=True))

    def test_journal_created_with_header(self):
        self.service.assign(self.ivan)
        self.assertEqual(self.rows()[0],
                         ("Группа", "Фамилия", "Имя", "Номер билета", "Дата и время", "Повтор (да/нет)"))

    def test_idempotent_same_ticket_new_row(self):
        first = self.service.assign(self.ivan)
        before = self.rows()
        second = self.service.assign(self.ivan)
        third = self.service.assign(Student("a-1", " иванов ", "ИВАН"))  # регистр/пробелы
        self.assertEqual(first.ticket.number, second.ticket.number)
        self.assertEqual(first.ticket.number, third.ticket.number)
        self.assertEqual([first.repeat, second.repeat, third.repeat], [False, True, True])
        rows = self.rows()
        self.assertEqual(len(rows), 4)
        self.assertEqual(rows[:2], before)  # старые строки не тронуты
        self.assertEqual([r[5] for r in rows[1:]], ["нет", "да", "да"])
        self.assertEqual({r[3] for r in rows[1:]}, {first.ticket.number})

    def test_same_name_in_other_group_is_a_different_student(self):
        self.service.assign(self.ivan)
        other = Student("B-2", "Иванов", "Иван")
        self.assertIsNone(self.journal.first_ticket_for(other))
        self.assertFalse(self.service.assign(other).repeat)

    def test_history_is_chronological(self):
        for student in (self.ivan, Student("A-1", "Петров", "Пётр"), self.ivan):
            self.service.assign(student)
        rows = self.rows()[1:]
        self.assertEqual([r[1] for r in rows], ["Иванов", "Петров", "Иванов"])
        stamps = [r[4] for r in rows]
        self.assertTrue(all(isinstance(s, datetime) for s in stamps))
        self.assertEqual(stamps, sorted(stamps))

    def test_ticket_is_one_of_parsed_numbers(self):
        for i in range(30):
            number = self.service.assign(Student("A-1", f"Ф{i}", "И")).ticket.number
            self.assertIn(number, range(1, 6))

    def test_locked_journal_raises_and_keeps_old_data(self):
        self.service.assign(self.ivan)
        before = self.rows()
        with mock.patch.object(journal_module.os, "replace", side_effect=PermissionError):
            with self.assertRaises(JournalLockedError):
                self.service.assign(Student("A-1", "Петров", "Пётр"))
        self.assertEqual(self.rows(), before)
        self.assertFalse(self.path.with_name("results.xlsx.tmp").exists())

    def test_ticket_missing_from_docx_raises(self):
        self.journal.append(self.ivan, 99, False)
        with self.assertRaises(TicketNotFoundError):
            self.service.assign(self.ivan)
        self.assertEqual(len(self.rows()), 2)  # ничего лишнего не дописано


if __name__ == "__main__":
    unittest.main()
