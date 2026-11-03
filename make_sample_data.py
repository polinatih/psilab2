"""Создаёт демонстрационные students.xlsx и tickets.docx рядом со скриптом."""
from pathlib import Path

from docx import Document
from openpyxl import Workbook

BASE = Path(__file__).resolve().parent

wb = Workbook()
data = {
    "ИС-21": [("Иванов", "Иван"), ("Петров", "Пётр"), ("Сидорова", "Анна")],
    "ИС-22": [("Козлов", "Дмитрий"), ("Смирнова", "Ольга")],
    "ПИ-23": [],  # пустая группа — для проверки
}
wb.remove(wb.active)
for group, students in data.items():
    ws = wb.create_sheet(group)
    ws.append(["Фамилия", "Имя"])
    for row in students:
        ws.append(list(row))
wb.save(BASE / "students.xlsx")

doc = Document()
for n in range(1, 8):
    doc.add_paragraph(f"Билет {n}")
    for q in range(1, 4):
        doc.add_paragraph(f"{q}. Вопрос {q} билета {n}")
    doc.add_paragraph("")
doc.save(BASE / "tickets.docx")
print("Созданы students.xlsx и tickets.docx")
