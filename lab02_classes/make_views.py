#!/usr/bin/env python3
"""Генерує з єдиного джерела classes.puml:
  * classes_core.puml — дидактичний фрагмент з усіма пʼятьма типами відношень;
  * classes_view1..4.puml — чотири часткові види повної діаграми.
Склад атрибутів і методів кожного класу береться дослівно з classes.puml, тому
розійтися з повною діаграмою він не може. Кожне з 30 відношень повної діаграми
потрапляє рівно на один із чотирьох видів.
"""
import re
import subprocess
from pathlib import Path

HERE = Path(__file__).resolve().parent
SRC = HERE / "classes.puml"

HEADER = '''@startuml
' {title}
' {note}
' Камінський Олексій Дмитрович, група ВТ-23-2

skinparam defaultFontName "Helvetica"
skinparam defaultFontSize 18
skinparam classAttributeIconSize 0
skinparam shadowing false
skinparam ArrowFontSize 15
skinparam nodesep 25
skinparam ranksep 40
skinparam class {{
  BackgroundColor #FEFEFE
  BorderColor #33668E
  ArrowColor #33668E
}}

'''

NOTE_VIEW = ("Частковий вид повної діаграми класів (classes.puml). Склад "
             "атрибутів і методів\n' кожного класу дослівно збігається з повною "
             "діаграмою.")
NOTE_CORE = ("Фрагмент повної діаграми класів (classes.puml) для демонстрації "
             "всіх пʼяти\n' типів відношень UML. Склад атрибутів і методів "
             "кожного класу дослівно\n' збігається з повною діаграмою.")

VIEWS = [
    dict(
        file="classes_core.puml",
        title="Лабораторна робота №2. Фрагмент діаграми класів: усі пʼять типів відношень",
        note=NOTE_CORE,
        classes=["Читач", "Публікація", "Книга", "Метадані", "Колекція",
                 "ФайлКниги", "IПарсерФормату", "ПарсерEPUB"],
        rels="""\' --- 1. Узагальнення (generalization) ---
Публікація <|-- Книга

\' --- 2. Реалізація інтерфейсу (realization) ---
ПарсерEPUB ..|> IПарсерФормату

\' --- 3. Композиція (composition) ---
Публікація "1" *-- "1" Метадані : описується
Публікація "1" *-- "1..*" ФайлКниги : містить
Читач "1" *-- "0..*" Колекція : веде

\' --- 4. Агрегація (aggregation) ---
Колекція "0..*" o-- "0..*" Публікація : містить

\' --- 5. Асоціація з множинністю (association) ---
ФайлКниги "0..*" -- "1" IПарсерФормату : обробляється >

\' --- підказки розкладки (не відношення UML) ---
Метадані -[hidden]down- ФайлКниги
""",
    ),
    dict(
        file="classes_view1.puml",
        title="Вид 1. Користувачі, доступ і модерація",
        classes=["КористувачСистеми", "Читач", "МодераторКонтенту",
                 "АдміністраторСистеми", "Сповіщення", "Скарга",
                 "ЗавданняМодерації", "Публікація"],
        rels="""' --- Узагальнення (3) ---
КористувачСистеми <|-- Читач
КористувачСистеми <|-- МодераторКонтенту
КористувачСистеми <|-- АдміністраторСистеми

' --- Агрегація (1) ---
ЗавданняМодерації "0..1" o-- "0..*" Скарга : групує

' --- Асоціації з множинністю (5) ---
КористувачСистеми "0..1" -- "0..*" Сповіщення : отримує >
Читач "1" -- "0..*" Скарга : подає >
МодераторКонтенту "0..1" -- "0..*" ЗавданняМодерації : виконує >
Скарга "0..*" -- "1" Публікація : стосується >
ЗавданняМодерації "0..*" -- "1" Публікація : перевіряє >
""",
    ),
    dict(
        file="classes_view2.puml",
        title="Вид 2. Контент: публікації, метадані, файли й колекції",
        classes=["Читач", "Публікація", "Книга", "Документ", "Метадані",
                 "ФайлКниги", "Каталог", "Колекція"],
        rels="""' --- Узагальнення (2) ---
Публікація <|-- Книга
Публікація <|-- Документ

' --- Композиція (4) ---
Публікація "1" *-- "1" Метадані : описується
Публікація "1" *-- "1..*" ФайлКниги : містить
Читач "1" *-- "0..*" Документ : володіє
Читач "1" *-- "0..*" Колекція : веде

' --- Агрегація (2) ---
Каталог "0..1" o-- "0..*" Книга : публікує
Колекція "0..*" o-- "0..*" Публікація : містить
""",
    ),
    dict(
        file="classes_view3.puml",
        title="Вид 3. Читання: прогрес і позначки",
        classes=["Читач", "Публікація", "ПрогресЧитання", "Позначка",
                 "Закладка", "Нотатка"],
        rels="""' --- Узагальнення (2) ---
Позначка <|-- Закладка
Позначка <|-- Нотатка

' --- Композиція (2) ---
Читач "1" *-- "0..*" ПрогресЧитання : накопичує
Читач "1" *-- "0..*" Позначка : створює

' --- Асоціації з множинністю (2) ---
ПрогресЧитання "0..*" -- "1" Публікація : фіксує позицію в >
Позначка "0..*" -- "1" Публікація : посилається на >
""",
    ),
    dict(
        file="classes_view4.puml",
        title="Вид 4. Підписка, оплата та обробка форматів",
        classes=["Читач", "АдміністраторСистеми", "Підписка", "ТарифнийПлан",
                 "Платіж", "ФайлКниги", "IПарсерФормату", "ПарсерEPUB",
                 "ПарсерPDF"],
        rels="""' --- Реалізація інтерфейсу (2) ---
ПарсерEPUB ..|> IПарсерФормату
ПарсерPDF ..|> IПарсерФормату

' --- Асоціації з множинністю (5) ---
Читач "0..1" -- "0..*" Підписка : оформлює >
Підписка "0..*" -- "1" ТарифнийПлан : діє за >
Підписка "0..1" -- "0..*" Платіж : оплачується >
АдміністраторСистеми "1" -- "0..*" ТарифнийПлан : керує >
ФайлКниги "0..*" -- "1" IПарсерФормату : обробляється >
""",
    ),
]


def extract_blocks(text: str) -> dict:
    """Повертає {назва: повний текст блоку} для кожного class/abstract class/interface."""
    out = {}
    pattern = re.compile(
        r"^[ \t]*((?:abstract\s+class|interface|class)\s+([^\s{]+)\s*\{)",
        re.MULTILINE)
    for m in pattern.finditer(text):
        name = m.group(2)
        start = m.start(1)
        depth = 0
        i = text.index("{", start)
        depth = 1
        j = i + 1
        while depth and j < len(text):
            if text[j] == "{":
                depth += 1
            elif text[j] == "}":
                depth -= 1
            j += 1
        block = text[start:j]
        # знімаємо спільний відступ пакета
        lines = block.split("\n")
        dedented = [lines[0].lstrip()] + [l[2:] if l.startswith("  ") else l
                                          for l in lines[1:]]
        out[name] = "\n".join(dedented)
    return out


REL_RE = re.compile(r'^(?!\s*\')(.*?)(<\|--|\.\.\|>|\*--|o--|--)(.*)$')


def relations(text: str) -> list:
    """Нормалізований перелік відношень (без прихованих підказок розкладки)."""
    out = []
    for line in text.split("\n"):
        line = line.strip()
        if not line or line.startswith("'") or "hidden" in line:
            continue
        if any(t in line for t in ("<|--", "..|>", "*--", "o--", '" -- "')):
            out.append(re.sub(r"\s+", " ", line))
    return out


def verify(src: str) -> None:
    """Перевіряє, що чотири види разом містять усі відношення повної діаграми
    рівно по одному разу, а класи — з тим самим складом членів."""
    full = relations(src)
    covered = []
    for view in VIEWS:
        if view["file"] == "classes_core.puml":
            continue
        covered += relations((HERE / view["file"]).read_text(encoding="utf-8"))
    missing = sorted(set(full) - set(covered))
    extra = sorted(set(covered) - set(full))
    dupes = sorted({r for r in covered if covered.count(r) > 1})
    print(f"відношень у повній діаграмі: {len(full)}; на чотирьох видах: {len(covered)}")
    if missing or extra or dupes:
        raise SystemExit(f"розбіжність!\n не покрито: {missing}\n зайві: {extra}\n дублі: {dupes}")
    print("перевірка покриття відношень: OK (кожне відношення рівно на одному виді)")


def main() -> None:
    src = SRC.read_text(encoding="utf-8")
    blocks = extract_blocks(src)
    for view in VIEWS:
        missing = [c for c in view["classes"] if c not in blocks]
        if missing:
            raise SystemExit(f"не знайдено класів у classes.puml: {missing}")
        body = HEADER.format(title=view["title"], note=view.get("note", NOTE_VIEW))
        body += "\n\n".join(blocks[c] for c in view["classes"])
        body += "\n\n" + view["rels"] + "@enduml\n"
        path = HERE / view["file"]
        path.write_text(body, encoding="utf-8")
        print("записано", path.name)

    verify(src)

    names = [v["file"] for v in VIEWS]
    subprocess.run(["plantuml", "-checkonly", *names], cwd=HERE, check=True)
    subprocess.run(["plantuml", "-tsvg", *names], cwd=HERE, check=True)
    subprocess.run(["plantuml", "-tpng", *names], cwd=HERE, check=True)
    for v in VIEWS:
        stem = v["file"][:-5]
        for ext in ("svg", "png"):
            f = HERE / f"{stem}.{ext}"
            print(f"{f.name}: {f.stat().st_size // 1024} КБ")


if __name__ == "__main__":
    main()
