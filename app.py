"""Лаб. 2: генератор билетов с графическим интерфейсом (Tkinter).

Зависимости: pip install openpyxl python-docx
Запуск:      python app.py   (students.xlsx и tickets.docx лежат рядом)
"""
import logging
import tkinter as tk
from pathlib import Path
from tkinter import font as tkfont
from tkinter import messagebox, ttk

from journal import ExcelJournal
from models import JournalError, JournalLockedError, SourceError, TicketNotFoundError
from service import TicketService
from sources import load_students, load_tickets

BASE_DIR = Path(__file__).resolve().parent
STUDENTS_FILE = BASE_DIR / "students.xlsx"
TICKETS_FILE = BASE_DIR / "tickets.docx"
RESULTS_FILE = BASE_DIR / "results.xlsx"
LOG_FILE = BASE_DIR / "lab2.log"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Генератор билетов")
        self.resizable(False, False)
        self.groups = {}
        self.service = None
        self._students = []

        big = tkfont.nametofont("TkDefaultFont").copy()
        big.configure(size=16, weight="bold")
        self._big_font = big

        frame = ttk.Frame(self, padding=16)
        frame.grid()
        ttk.Label(frame, text="Группа").grid(row=0, column=0, sticky="w")
        self.group_cb = ttk.Combobox(frame, state="readonly", width=36)
        self.group_cb.grid(row=1, column=0, pady=(0, 10))
        ttk.Label(frame, text="Студент").grid(row=2, column=0, sticky="w")
        self.student_cb = ttk.Combobox(frame, state="readonly", width=36)
        self.student_cb.grid(row=3, column=0, pady=(0, 14))
        self.generate_btn = ttk.Button(frame, text="Сгенерировать билет",
                                       state="disabled", command=self.generate)
        self.generate_btn.grid(row=4, column=0, sticky="ew")
        ttk.Button(frame, text="Перечитать файлы", command=self.reload).grid(
            row=5, column=0, sticky="ew", pady=(6, 0))
        self.status = ttk.Label(frame, text="", wraplength=300, foreground="gray40")
        self.status.grid(row=6, column=0, sticky="w", pady=(10, 0))

        self.group_cb.bind("<<ComboboxSelected>>", self._on_group)
        self.student_cb.bind("<<ComboboxSelected>>", lambda e: self._update_button())
        self.reload()

    # ---------- загрузка данных ----------
    def reload(self):
        self.groups, self.service = {}, None
        errors = []
        try:
            self.groups = load_students(STUDENTS_FILE)
        except SourceError as exc:
            errors.append(str(exc))
        tickets = []
        try:
            tickets = load_tickets(TICKETS_FILE)
        except SourceError as exc:
            errors.append(str(exc))

        self.group_cb.set("")
        self._fill_students([])
        if errors:
            self.group_cb["values"] = ()
            self.status.configure(text="\n".join(errors), foreground="firebrick")
            messagebox.showerror("Ошибка входных данных", "\n\n".join(errors))
            return
        self.service = TicketService(ExcelJournal(RESULTS_FILE), tickets)
        self.group_cb["values"] = list(self.groups)
        self.status.configure(
            text=f"Групп: {len(self.groups)}, билетов: {self.service.tickets_count}",
            foreground="gray40")

    # ---------- выбор ----------
    def _fill_students(self, students):
        self._students = students
        self.student_cb["values"] = [s.full_name for s in students]
        self.student_cb.set("")
        self._update_button()

    def _on_group(self, _event=None):
        self._fill_students(self.groups.get(self.group_cb.get(), []))

    def _update_button(self):
        ready = self.service is not None and self.student_cb.current() >= 0
        self.generate_btn.configure(state="normal" if ready else "disabled")

    # ---------- генерация ----------
    def generate(self):
        index = self.student_cb.current()
        if self.service is None or index < 0:
            return
        student = self._students[index]
        while True:
            try:
                assignment = self.service.assign(student)
                break
            except JournalLockedError:
                if not messagebox.askretrycancel(
                        "Файл занят",
                        "Не удалось записать results.xlsx: файл открыт в другой "
                        "программе (например, в Excel).\nЗакройте его и нажмите «Повторить»."):
                    return
            except (JournalError, TicketNotFoundError) as exc:
                messagebox.showerror("Ошибка", str(exc))
                return
        self._show_result(student, assignment)
        self.student_cb.set("")  # готово к следующему студенту, группа остаётся
        self._update_button()

    def _show_result(self, student, assignment):
        win = tk.Toplevel(self)
        win.title("Ваш билет")
        win.resizable(False, False)
        win.transient(self)
        frame = ttk.Frame(win, padding=20)
        frame.grid()
        ttk.Label(frame, text=f"Ваш билет № {assignment.ticket.number}",
                  font=self._big_font).grid(sticky="w")
        ttk.Label(frame, text=f"{student.group}, {student.full_name}",
                  foreground="gray40").grid(sticky="w", pady=(0, 10))
        for i, question in enumerate(assignment.ticket.questions, 1):
            ttk.Label(frame, text=f"{i}. {question}", wraplength=460,
                      justify="left").grid(sticky="w", pady=3)
        if assignment.repeat:
            ttk.Label(frame, text="Повторный выбор: билет остался прежним.",
                      foreground="gray40").grid(sticky="w", pady=(10, 0))
        ttk.Button(frame, text="Закрыть (Esc)", command=win.destroy).grid(
            sticky="e", pady=(14, 0))
        win.bind("<Escape>", lambda e: win.destroy())
        win.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - win.winfo_width()) // 2
        y = self.winfo_rooty() + 30
        win.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        win.grab_set()
        win.focus_force()
        self.wait_window(win)


def main():
    logging.basicConfig(filename=LOG_FILE, level=logging.INFO, encoding="utf-8",
                        format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    App().mainloop()


if __name__ == "__main__":
    main()
