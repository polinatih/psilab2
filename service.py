"""Бизнес-логика назначения билета. Не знает ни про Excel, ни про UI."""
import random
from typing import Optional, Protocol, Sequence

from models import Assignment, Student, Ticket, TicketNotFoundError


class Journal(Protocol):
    def first_ticket_for(self, student: Student) -> Optional[int]: ...
    def append(self, student: Student, ticket_number: int, repeat: bool) -> None: ...


class TicketService:
    def __init__(self, journal: Journal, tickets: Sequence[Ticket], rng=None):
        if not tickets:
            raise ValueError("Список билетов пуст")
        self._journal = journal
        self._tickets = {t.number: t for t in tickets}
        self._numbers = sorted(self._tickets)
        self._rng = rng or random.Random()

    @property
    def tickets_count(self) -> int:
        return len(self._tickets)

    def assign(self, student: Student) -> Assignment:
        first = self._journal.first_ticket_for(student)
        if first is not None:
            ticket = self._tickets.get(first)
            if ticket is None:
                raise TicketNotFoundError(
                    f"Студенту уже выдан билет № {first}, но его нет в tickets.docx")
            repeat = True
        else:
            # выбираем среди реально распознанных билетов (пропущенные не выдаём)
            ticket = self._tickets[self._rng.choice(self._numbers)]
            repeat = False
        self._journal.append(student, ticket.number, repeat)  # сохранение сразу
        return Assignment(ticket, repeat)
