# -*- coding: utf-8 -*-
"""
Модель данных модуля «Ремонт».

Реестр ремонтов хранится видимым листом «Ремонт» в ПНОС_ТТП.xlsm (единственное
хранилище модуля — файлы остаются обычными, состояние не прячется). Эти
dataclass-ы — представление одной строки реестра в памяти.
"""

from __future__ import annotations

from dataclasses import dataclass, field


# Статус исполнительной документации по записи реестра.
ID_PRESENT = "есть"        # ИД поступила → строка «Сведения» состояния С2/С3
ID_EXPECTED = "ожидается"  # только акты, ИД ждём → состояние С1
ID_NEVER = "не будет"      # ИД не будет → состояние С1 постоянно


@dataclass
class ActRef:
    """Ссылка на акт ревизии и отбраковки."""
    number: str            # напр. «222-1-84»
    date: str = ""         # ДД.ММ.ГГГГ, из грифа УТВЕРЖДАЮ самого акта

    def is_valid(self) -> bool:
        return bool(self.number.strip())


@dataclass
class RepairRecord:
    """
    Одна запись реестра ремонтов = один комплект исполнительной документации
    (или акт без ИД). Один комплект ИД может закрывать несколько актов.
    """
    id_status: str = ID_EXPECTED     # ID_PRESENT / ID_EXPECTED / ID_NEVER
    date_done: str = ""              # дата выполнения работ («Сдал» в реестре ИД)
    acts: list[ActRef] = field(default_factory=list)
    replacements: str = ""           # типоразмеры из таблиц актов (для сверки)
    material: str = ""               # материал замены (показывается эксперту)
    source_files: list[str] = field(default_factory=list)
    note: str = ""                   # примечание сотрудника

    @property
    def has_id(self) -> bool:
        """ИД поступила и есть дата выполнения — строка идёт в состояние С2/С3."""
        return self.id_status == ID_PRESENT and bool(self.date_done.strip())


@dataclass
class Section:
    """
    Сечение = точка замера на карте и строка листа «Импорт» в Excel.
    Толщины хранятся строками, как в отчёте («10.0», «-», «12»), чтобы
    сравнение с эталоном было точным.
    """
    number: int                      # № сечения (колонка «Сечение»)
    element: str                     # тип: Труба / Отвод / Переход / Задвижка / Бобышка …
    thickness_nom: str               # исполнительная (номинальная) толщина, Тисп
    diameter: str                    # диаметр, мм (может быть двусторонним для перехода)
    measures: list[str] = field(default_factory=list)  # замеры т.1…т.6 (сколько заполнено)
    tmin: str = ""                   # минимальная выявленная


@dataclass
class Change:
    """Одно изменение для экрана предпросмотра «было → станет»."""
    target: str        # «Импорт» / «Word» / «Карта»
    where: str         # напр. «сечение 8» / «таблица 1»
    before: str
    after: str
    enabled: bool = True   # галочка сотрудника; можно исключить из применения
