from playwright.sync_api import sync_playwright

from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

from io import BytesIO

from datetime import datetime

import re
import time
import hashlib

from pypdf import PdfReader
from docx import Document
from openpyxl import load_workbook

import xlrd
from legacy_doc import extract_text as extract_legacy_doc_text
import olefile


# ============================================================
# НАСТРОЙКИ
# ============================================================

MAX_PRICE = "5000000"

EXPECTED_OKPD2_CODES = [
    "71.11",
    "71.11.1",
    "71.11.10",
    "71.11.10.000",

    "71.11.2",
    "71.11.21",
    "71.11.21.000",
    "71.11.22",
    "71.11.22.000",
    "71.11.23",
    "71.11.23.000",
    "71.11.24",
    "71.11.24.000",

    "71.11.3",
    "71.11.31",
    "71.11.32",
    "71.11.32.000",
    "71.11.33",
    "71.11.33.000",
    "71.11.33.100",
    "71.11.33.900",

    "71.11.4",
    "71.11.41",
    "71.11.41.000",
    "71.11.41.100",
    "71.11.41.200",
    "71.11.42",
]

EXCLUDED_OKPD2 = {
    "71.11.23",
    "71.11.23.000",
    "71.11.33.100",
}

TEST_TENDER_NUMBER = "0373200081226000359"

# ============================================================
# БЫСТРЫЙ ПРЕДВАРИТЕЛЬНЫЙ ФИЛЬТР ТЕНДЕРОВ
# ============================================================

# Жёсткие признаки явно тяжёлых / неподходящих закупок.
# Если встречаются уже в названии закупки — такую закупку
# можно не отправлять дальше на обработку документов.
QUICK_REJECT_TITLE_KEYWORDS = [

    # --------------------------------------------------------
    # Сметы / комплексное проектирование
    # --------------------------------------------------------

    "проектно-сметн",
    "проектно - сметн",
    "сметн",


    # --------------------------------------------------------
    # Инженерные изыскания
    # --------------------------------------------------------

    "инженерн изыск",
    "инженерные изыск",
    "инженерных изыск",
    "инженерно-изыскатель",
    "проектно-изыскатель",
    "проектно - изыскатель",
    "изыскательск работ",
    "изысканий",


    # --------------------------------------------------------
    # Кадастровые работы
    # --------------------------------------------------------

    "кадастров работ",
    "кадастровые работы",
    "кадастровых работ",


    # --------------------------------------------------------
    # Экспертиза
    # --------------------------------------------------------

    "государственн экспертиз",
    "положительн заключени экспертиз",
    "заключени экспертиз",
    "достоверност сметн",


    # --------------------------------------------------------
    # Обследования
    # --------------------------------------------------------

    "техническ обслед",
    "обследовани техническ состояни",
    "оценк техническ состояни",


    # --------------------------------------------------------
    # Надзор
    # --------------------------------------------------------

    "авторск надзор",
    "авторский надзор",


    # --------------------------------------------------------
    # Инженерные сети
    # --------------------------------------------------------

    "электроснабж",
    "электрическ сет",
    "электрическ лини",
    "кабельн лини",
    "кабельн сет",
    "внешн сет",
    "газоснабж",
    "водоснабж",
    "канализац",
    "теплоснабж",
    "тепловой пункт",
    "теплового пункта",


    # --------------------------------------------------------
    # Газопроводы
    # --------------------------------------------------------

    "газопровод",
    "магистральн газопровод",


    # --------------------------------------------------------
    # Дороги / транспортная инфраструктура
    # --------------------------------------------------------

    "строительств мост",
    "реконструкц мост",
    "путепровод",
    "автомобильн дорог",
    "автомобильной дорог",
    "автомобильные дорог",
    "автодорог",


    # --------------------------------------------------------
    # Земляные и физические работы
    # --------------------------------------------------------

    "отсыпк грунт",
    "отсыпк территории",
    "разработк грунта",
    "выемк грунта",
    "землян работ",


    # --------------------------------------------------------
    # Взрывоопасные предметы
    # --------------------------------------------------------

    "взрывоопасн",
    "взрывоопасных предмет",


    # --------------------------------------------------------
    # Узкоспециализированные работы
    # --------------------------------------------------------

    "обеспечению пожарной безопасност",


    # --------------------------------------------------------
    # Неархитектурные услуги
    # --------------------------------------------------------

    "инвентаризац кладбищ",
    "инвентаризация кладбищ",
]

# ============================================================
# СОЧЕТАНИЯ СЛОВ ДЛЯ БОЛЕЕ НАДЁЖНОГО ОТБОРА
# ============================================================

# Здесь условия проверяются отдельно:
# оба признака должны присутствовать в названии,
# но между ними могут быть любые слова.

QUICK_REJECT_TITLE_COMBINATIONS = [
    ("отсыпк", "грунт"),
    ("разработк грунт",),
    ("выемк грунт",),

    ("инвентаризац", "кладбищ"),
    ("захорон", "кладбищ"),

    ("землян", "работ"),
]

# Признаки интересного для нас типа работ.
# Они НЕ являются автоматическим "БРАТЬ".
# Они просто помогают понять, что закупку стоит передать ИИ.
QUICK_GOOD_TITLE_KEYWORDS = [
    "дизайн-проект",
    "дизайн проект",
    "благоустрой",
    "ландшафт",
    "эскиз",
    "концепц",
    "архитектурн",
    "планировк",
    "генеральн план",
    "организац земельн",
]

# Жёсткие признаки в названиях файлов.
# Если такие файлы есть, закупка может быть слишком тяжёлой.
QUICK_REJECT_FILE_KEYWORDS = [
    "проектно-смет",
    "сметн",
    "локальн смет",
    "сводн смет",
    "инженерн изыск",
    "инженерные изыск",
    "государственн экспертиз",
    "заключение экспертиз",
    "техническ обслед",
    "авторск надзор",
]

def quick_screen_title(title):
    """
    Быстрая проверка только по названию закупки.

    Проверяет:
    1. Обычные красные флаги.
    2. Сочетания нескольких признаков.
    3. Положительные признаки.

    Возвращает:
        {
            "status": "candidate" | "quick_reject",
            "good_matches": [...],
            "reject_matches": [...],
            "reason": "..."
        }
    """

    text = (title or "").lower().replace("ё", "е")

    reject_matches = []
    good_matches = []

    # --------------------------------------------------------
    # Обычные красные флаги
    # --------------------------------------------------------

    for keyword in QUICK_REJECT_TITLE_KEYWORDS:

        if keyword in text:
            reject_matches.append(keyword)

    # --------------------------------------------------------
    # Сочетания нескольких признаков
    # --------------------------------------------------------

    for combination in QUICK_REJECT_TITLE_COMBINATIONS:

        if all(part in text for part in combination):

            reject_matches.append(
                " + ".join(combination)
            )

    # --------------------------------------------------------
    # Положительные признаки
    # --------------------------------------------------------

    for keyword in QUICK_GOOD_TITLE_KEYWORDS:

        if keyword in text:
            good_matches.append(keyword)

    # --------------------------------------------------------
    # Итог
    # --------------------------------------------------------

    if reject_matches:

        return {
            "status": "quick_reject",
            "good_matches": good_matches,
            "reject_matches": reject_matches,
            "reason": (
                "В названии обнаружены признаки "
                "сложной/неподходящей закупки"
            ),
        }

    return {
        "status": "candidate",
        "good_matches": good_matches,
        "reject_matches": [],
        "reason": (
            "Явных причин для раннего отсева "
            "по названию не найдено"
        ),
    }


def quick_screen_files(file_list):
    """
    Проверка только названий файлов.

    Ничего из файлов ещё не скачивается.
    """

    reject_matches = []

    for file_info in file_list:

        filename = (
            file_info.get("name", "")
            .lower()
            .replace("ё", "е")
        )

        for keyword in QUICK_REJECT_FILE_KEYWORDS:

            if keyword in filename:

                reject_matches.append({
                    "file": file_info.get("name", ""),
                    "keyword": keyword,
                })

    if reject_matches:
        return {
            "status": "quick_reject",
            "reject_matches": reject_matches,
            "reason": (
                "В названиях документов обнаружены "
                "признаки сложной/неподходящей закупки"
            ),
        }

    return {
        "status": "candidate",
        "reject_matches": [],
        "reason": (
            "По названиям файлов явных "
            "красных флагов не найдено"
        ),
    }

# ============================================================
# ГЛУБОКИЙ ПРЕДВАРИТЕЛЬНЫЙ АНАЛИЗ ТЕКСТА ДОКУМЕНТОВ
# ============================================================

DEEP_RED_FLAGS = {
        "expertise": [
        # Подрядчик обязан получить положительное заключение
        r"подрядчик\w*.{0,180}"
        r"(?:получит\w*|получен\w*|обеспечит\w*|обеспеч\w*|предоставит\w*)"
        r".{0,180}"
        r"положительн\w*\s+заключени\w*"
        r".{0,100}"
        r"государственн\w*\s+экспертиз\w*",

        # Обязательства подрядчика зависят от получения экспертизы
        r"обязательств\w*\s+подрядчик\w*"
        r".{0,180}"
        r"счита\w*\s+выполнен\w*"
        r".{0,180}"
        r"получен\w*\s+.*заключени\w*"
        r".{0,100}"
        r"государственн\w*\s+экспертиз\w*",

        # Явно установлена обязанность получить заключение
        r"(?:необходим\w*|требует\w*|требован\w*|обязательн\w*)"
        r".{0,120}"
        r"(?:получит\w*|получен\w*|получени\w*)"
        r".{0,120}"
        r"(?:положительн\w*\s+)?заключени\w*"
        r".{0,100}"
        r"государственн\w*\s+экспертиз\w*",
    ],

    "engineering": [
        # Реальное выполнение инженерных изысканий
        r"(?:выполн\w*|провед\w*|осуществля\w*|выполнен\w*)"
        r".{0,120}"
        r"инженерн\w*[-\s]+геодезическ\w*\s+изыскан\w*",

        r"(?:выполн\w*|провед\w*|осуществля\w*|выполнен\w*)"
        r".{0,120}"
        r"инженерн\w*[-\s]+геологическ\w*\s+изыскан\w*",

        r"(?:выполн\w*|провед\w*|осуществля\w*|выполнен\w*)"
        r".{0,120}"
        r"инженерн\w*\s+изыскан\w*",

        # Явная обязанность подрядчика
        r"подрядчик\w*.{0,120}"
        r"(?:выполн\w*|провест\w*|осуществ\w*).{0,120}"
        r"инженерн\w*[-\s]+геодезическ\w*",

        r"подрядчик\w*.{0,120}"
        r"(?:выполн\w*|провест\w*|осуществ\w*).{0,120}"
        r"инженерн\w*[-\s]+геологическ\w*",
    ],

    "estimates": [
        r"подготов\w*\s+проектн\w*\s+и\s+сметн\w*\s+документаци",
        r"подготов\w*\s+сметн\w*\s+документаци",
        r"разработ\w*\s+сметн\w*\s+документаци",
        r"проектн\w*\s*[-–—]\s*сметн\w*",
        r"сметн\w*\s+стоимост\w*",
        r"положительн\w*\s+заключени\w*\s+.*сметн\w*",
    ],

    "working_documentation": [
        r"рабоч\w*\s+документаци",
        r"рабоч\w*\s+проект\w*",
        r"комплект\w*\s+рабоч\w*\s+чертеж\w*",
        r"раздел\w*\s+рабоч\w*\s+документаци",
    ],

    "specialists": [
        # Подрядчик/исполнитель обязан привлечь конкретных специалистов
        r"(?:подрядчик\w*|исполнитель\w*).{0,120}"
        r"(?:обязан\w*|должен\w*|требует\w*|необходимо\w*).{0,100}"
        r"(?:привлеч\w*|обеспеч\w*|иметь\w*|имет\w*).{0,100}"
        r"(?:геолог\w*|геодезист\w*|кадастров\w*|оценщик\w*|сметчик\w*)",

        # Наличие специалиста именно в штате / персонале
        r"(?:в\s+штате|в\s+составе\s+персонала).{0,100}"
        r"(?:геолог\w*|геодезист\w*|кадастров\w*|оценщик\w*|сметчик\w*)",

        # Явное требование специалиста определённого профиля
        r"(?:наличи\w*|наличие|требован\w*).{0,100}"
        r"(?:специалист\w*|сотрудник\w*|работник\w*).{0,100}"
        r"(?:геолог\w*|геодезист\w*|кадастров\w*|оценщик\w*|сметчик\w*)",
    ],

    "unlimited_revisions": [
        r"до\s+полного\s+согласован\w*",
        r"без\s+ограничени\w*\s+количеств\w*",
        r"устран\w*\s+всех\s+замечани\w*",
        r"без\s+ограничени\w*\s+числ\w*\s+замечани\w*",
        r"до\s+достижени\w*\s+результат\w*.*удовлетворяющ\w*\s+заказчик\w*",
    ],

    "approvals": [
        # Согласование результата с внешними организациями
        r"согласован\w*\s+проектн\w*\s+документаци\w*"
        r".{0,100}"
        r"(?:государственн\w*|администраци\w*|министерств\w*|ведомств\w*|сетев\w*|ресурсоснабжающ\w*)",

        r"получен\w*\s+разрешен\w*"
        r".{0,100}"
        r"(?:строительств\w*|реконструкци\w*|эксплуат\w*)",

        r"получен\w*\s+согласован\w*"
        r".{0,100}"
        r"(?:администраци\w*|ведомств\w*|сетев\w*|ресурсоснабжающ\w*)",
    ],

    "supervision": [
        r"авторск\w*\s+надзор",
        r"техническ\w*\s+надзор",
        r"строительн\w*\s+контрол\w*",
    ],

    "construction": [
        r"выполн\w*\s+строительн\w*\s+работ",
        r"производств\w*\s+строительн\w*\s+работ",
        r"монтажн\w*\s+работ",
        r"демонтажн\w*\s+работ",
        r"землян\w*\s+работ",
        r"разработк\w*\s+грунт\w*",
        r"выемк\w*\s+грунт\w*",
        r"отсыпк\w*\s+грунт\w*",
    ],
}


DEEP_GOOD_FLAGS = {

    "concept": [
        r"дизайн[-\s]?проект",
        r"концепци[яи]",
        r"концептуальн\w+",
        r"эскизн\w+\s+проект",
        r"эскизн\w+\s+решени",
        r"архитектурн\w+\s+концепци",
    ],

    "landscape": [
        r"ландшафтн\w+",
        r"благоустройств\w+",
        r"озеленени[ея]",
        r"малые\s+архитектурные\s+формы",
    ],

    "visualization": [
        r"визуализаци[яи]",
        r"3d[-\s]?визуализаци",
        r"фотореалистичн\w+",
        r"рендер",
    ],

    "architectural_graphics": [
        # Реальная разработка графических материалов
        r"разработ\w*.{0,100}графическ\w+\s+материал",
        r"подготов\w*.{0,100}графическ\w+\s+материал",

        # Чертежи как результат проектных работ
        r"разработ\w*.{0,100}чертеж\w*",
        r"подготов\w*.{0,100}чертеж\w*",

        # Эскизный / концептуальный результат
        r"эскизн\w*\s+проект\w*",
        r"разработ\w*.{0,100}эскиз\w*",

        # Визуализация
        r"разработ\w*.{0,100}визуализац\w*",
        r"подготов\w*.{0,100}визуализац\w*",
    ],
}


def find_text_matches(text, patterns):
    """
    Ищет признаки в тексте документа.

    Возвращает список найденных выражений.
    """

    normalized = (
        text or ""
    ).lower().replace("ё", "е")

    matches = []

    for pattern in patterns:

        try:

            if re.search(
                pattern,
                normalized,
                flags=re.IGNORECASE
            ):

                matches.append(pattern)

        except Exception:
            continue

    return matches


def get_text_snippets(text, pattern, max_snippets=2):
    snippets = []

    try:
        matches = list(re.finditer(pattern, text, re.IGNORECASE))
    except re.error:
        return []

    for match in matches:
        start = max(0, match.start() - 250)
        end = min(len(text), match.end() + 350)

        snippet = text[start:end].replace("\n", " ").strip()

        # Убираем слишком короткие и дублирующиеся фрагменты
        if len(snippet) < 30:
            continue

        if snippet in snippets:
            continue

        snippets.append(snippet)

        if len(snippets) >= max_snippets:
            break

    return snippets


def deep_screen_documents(
    documents_text,
    loaded_files=None
):
    """
    Вторичный анализ уже извлечённого текста документов.

    ВАЖНО:
    Эта функция НЕ принимает окончательное решение.
    Она только собирает признаки для дальнейшего анализа.
    """

    text = documents_text or ""

    red_matches = {}
    good_matches = {}

    # --------------------------------------------------------
    # Красные флаги
    # --------------------------------------------------------

    for category, patterns in DEEP_RED_FLAGS.items():

        found = find_text_matches(
            text,
            patterns
        )

        if found:

            red_matches[category] = {
                "count": len(found),
                "patterns": found,
                "snippets": []
            }

            for pattern in found:

                snippets = get_text_snippets(
                    text,
                    pattern,
                    max_snippets=2
                )

                red_matches[category][
                    "snippets"
                ].extend(
                    snippets
                )


    # --------------------------------------------------------
    # Положительные признаки
    # --------------------------------------------------------

    for category, patterns in DEEP_GOOD_FLAGS.items():

        found = find_text_matches(
            text,
            patterns
        )

        if found:

            good_matches[category] = {
                "count": len(found),
                "patterns": found,
                "snippets": []
            }

            for pattern in found:

                snippets = get_text_snippets(
                    text,
                    pattern,
                    max_snippets=2
                )

                good_matches[category][
                    "snippets"
                ].extend(
                    snippets
                )


    # --------------------------------------------------------
    # Если есть отдельные файлы,
    # запоминаем, в каких именно файлах найдены признаки.
    # --------------------------------------------------------

    file_findings = []

    if loaded_files:

        for file_data in loaded_files:

            file_name = file_data.get(
                "name",
                "Без названия"
            )

            file_text = file_data.get(
                "text",
                ""
            )

            if not file_text:
                continue

            file_red = []

            for category, patterns in DEEP_RED_FLAGS.items():

                if find_text_matches(
                    file_text,
                    patterns
                ):

                    file_red.append(
                        category
                    )

            file_good = []

            for category, patterns in DEEP_GOOD_FLAGS.items():

                if find_text_matches(
                    file_text,
                    patterns
                ):

                    file_good.append(
                        category
                    )

            if file_red or file_good:

                file_findings.append({
                    "file": file_name,
                    "red_flags": file_red,
                    "good_flags": file_good,
                })


    return {
        "red_flags": red_matches,
        "good_flags": good_matches,
        "file_findings": file_findings,
    }

def extract_evidence_snippets(text, patterns, max_snippets=3):
    """
    Ищет небольшие фрагменты текста вокруг важных условий.
    Используется для структурного анализа закупки.
    """
    snippets = []

    for pattern in patterns:
        try:
            matches = re.finditer(pattern, text, re.IGNORECASE)
        except re.error:
            continue

        for match in matches:
            start = max(0, match.start() - 250)
            end = min(len(text), match.end() + 500)

            snippet = text[start:end]
            snippet = re.sub(r"\s+", " ", snippet).strip()

            if len(snippet) < 30:
                continue

            if snippet in snippets:
                continue

            snippets.append(snippet)

            if len(snippets) >= max_snippets:
                return snippets

    return snippets


def build_work_profile(
    documents_text,
    deep_analysis,
    loaded_files=None,
    title=""
):
    """
    Структурированный анализ закупки.

    Основные правила:
    1. Тип работы определяем прежде всего из названия и ТЗ.
    2. Не считаем случайное упоминание слова в документе видом работы.
    3. Риски engineering/expertise/etc. берём из уже проверенного deep_analysis.
    4. Бесплатные доработки и неограниченные доработки разделяем.
    """

    documents_text = documents_text or ""
    loaded_files = loaded_files or []
    title = title or ""

    free_revision_patterns = [
    r"без\s+дополнительной\s+оплаты.{0,300}(?:изменен|дополнен|доработ|коррект)",
    r"(?:изменен|дополнен|доработ|коррект).{0,300}без\s+дополнительной\s+оплаты",
    r"(?:изменен|дополнен|доработ).{0,300}бесплатн",
    r"доработан\w*\s+по\s+замечани\w*.{0,300}бесплатн",
    r"вносит\s+в\s+проект\w*\s+изменен\w*\s+и\s+дополнен\w*.{0,300}бесплатн",
    r"вносить\s+.*?изменени\w*\s+и\s+дополнени\w*.{0,300}бесплатн",
    ]
    
    profile = {
        "work_types": [],
        "deliverables": [],

        "completion_terms": [],
        "revision_terms": [],
        "approval_terms": [],

        "risks": {
            "engineering": False,
            "specialists": False,
            "estimates": False,
            "expertise": False,
            "approvals": False,
            "unlimited_revisions": False,
            "free_revisions": False,
            "supervision": False,
            "construction": False,
        },

        "positive_signals": {
            "concept": False,
            "landscape": False,
            "visualization": False,
            "architectural_graphics": False,
        },
    }

    free_revisions_found = False

    for pattern in free_revision_patterns:
        if re.search(pattern, documents_text, re.IGNORECASE | re.DOTALL):
            free_revisions_found = True
            break

    if free_revisions_found:
        profile["risks"]["free_revisions"] = True

    # ========================================================
    # ПОЛУЧАЕМ ТЕКСТ РАБОЧИХ ДОКУМЕНТОВ
    # ========================================================

    work_text_parts = []
    contract_text_parts = []

    for file_info in loaded_files:

        name = (
            file_info.get("name")
            or ""
        ).lower()

        text = (
            file_info.get("text")
            or ""
        )

        if not text.strip():
            continue

        # ТЗ / описание объекта
        if any(
            keyword in name
            for keyword in [
                "техническое задание",
                "техническое_задание",
                "тз ",
                "тз.",
                "описание объекта",
                "описание объекта закупки",
                "описание_объекта",
            ]
        ):

            work_text_parts.append(
                text
            )

        # Проект контракта
        if any(
            keyword in name
            for keyword in [
                "проект контракта",
                "проект государственного контракта",
                "проект мк",
                "проект государственного договора",
                "контракт",
                "договор",
            ]
        ):

            contract_text_parts.append(
                text
            )


    work_text = (
        "\n".join(work_text_parts)
        if work_text_parts
        else documents_text
    )

    contract_text = (
        "\n".join(contract_text_parts)
        if contract_text_parts
        else documents_text
    )


    # ========================================================
    # УЖЕ ПОДТВЕРЖДЁННЫЕ РИСКИ
    # ========================================================

    red_flags = (
        deep_analysis.get(
            "red_flags",
            {}
        )
        if deep_analysis
        else {}
    )

    good_flags = (
        deep_analysis.get(
            "good_flags",
            {}
        )
        if deep_analysis
        else {}
    )

    for key in [
        "engineering",
        "specialists",
        "estimates",
        "expertise",
        "supervision",
        "construction",
    ]:

        profile["risks"][key] = (
            key in red_flags
        )


    for key in profile["positive_signals"]:

        profile["positive_signals"][key] = (
            key in good_flags
        )


    # ========================================================
    # ТЕКСТ ДЛЯ ОПРЕДЕЛЕНИЯ СОСТАВА РАБОТ
    # ========================================================

    # Начинаем с названия — это самый надёжный источник
    focused_parts = [title]

    # Добавляем только строки ТЗ,
    # где одновременно есть действие + предмет.
    action_pattern = re.compile(
        r"(разработ\w*|"
        r"подготов\w*|"
        r"выполн\w*|"
        r"состав\w*|"
        r"создан\w*|"
        r"предостав\w*|"
        r"переда\w*|"
        r"оформ\w*|"
        r"сформир\w*|"
        r"согласован\w*)",
        re.IGNORECASE
    )

    domain_pattern = re.compile(
        r"(проект\w*|"
        r"документаци\w*|"
        r"чертеж\w*|"
        r"графическ\w*|"
        r"эскиз\w*|"
        r"визуализац\w*|"
        r"благоустрой\w*|"
        r"ландшафт\w*|"
        r"озеленени\w*|"
        r"планировк\w*|"
        r"межевани\w*|"
        r"генеральн\w*\s+план\w*)",
        re.IGNORECASE
    )

    for line in re.split(
        r"[\r\n]+",
        work_text
    ):

        line = re.sub(
            r"\s+",
            " ",
            line
        ).strip()

        if len(line) < 20:
            continue

        if (
            action_pattern.search(line)
            and domain_pattern.search(line)
        ):

            focused_parts.append(
                line
            )


    focus_text = "\n".join(
        focused_parts
    )


    # ========================================================
    # ТИПЫ РАБОТ
    # ========================================================

    # Название имеет приоритет.
    title_lower = title.lower()

        # --------------------------------------------------------
    # Градостроительная документация
    # --------------------------------------------------------

    if (
        "правил землепользования и застройки" in title_lower
        or "пзз" in title_lower
    ):

        profile["work_types"].append(
            "правила землепользования и застройки"
        )


    if (
        "генеральный план" in title_lower
        or "генерального плана" in title_lower
        or "генеральным планом" in title_lower
    ):

        profile["work_types"].append(
            "генеральный план"
        )


    if (
        "градостроительного плана" in title_lower
        or "градостроительный план" in title_lower
    ):

        profile["work_types"].append(
            "градостроительный план"
        )


    # --------------------------------------------------------
    # Кадастровые / границы
    # --------------------------------------------------------

    if (
        "кадастров" in title_lower
        or "границ территориальных зон" in title_lower
        or "границ населенных пунктов" in title_lower
        or "егрн" in title_lower
    ):

        profile["work_types"].append(
            "кадастровые / территориальные работы"
        )

    if (
        "дизайн-проект" in title_lower
        or "дизайн проект" in title_lower
    ):

        profile["work_types"].append(
            "дизайн-проект"
        )


    if (
        "благоустрой" in title_lower
        or "ландшафт" in title_lower
        or "озеленени" in title_lower
    ):

        profile["work_types"].append(
            "ландшафт"
        )


    if "планировк" in title_lower:

        profile["work_types"].append(
            "планировка территории"
        )


    if "межевани" in title_lower:

        profile["work_types"].append(
            "межевание"
        )


    if "визуализац" in title_lower:

        profile["work_types"].append(
            "визуализация"
        )


    # Подтверждение из ТЗ
    if (
        "проектирование" not in profile["work_types"]
        and re.search(
            r"(разработ\w*|подготов\w*).{0,100}"
            r"проектн\w+\s+документаци",
            focus_text,
            re.IGNORECASE
        )
    ):

        profile["work_types"].append(
            "проектирование"
        )


    if (
        "планировка территории"
        not in profile["work_types"]
        and re.search(
            r"проект\w*\s+планировк\w*",
            focus_text,
            re.IGNORECASE
        )
    ):

        profile["work_types"].append(
            "планировка территории"
        )


    if (
        "межевание"
        not in profile["work_types"]
        and re.search(
            r"проект\w*.{0,80}межевани\w*",
            focus_text,
            re.IGNORECASE
        )
    ):

        profile["work_types"].append(
            "межевание"
        )


    if (
        "ландшафт"
        not in profile["work_types"]
        and re.search(
            r"(?:разработ\w*|подготов\w*|выполн\w*).{0,120}"
            r"(?:ландшафтн\w*|благоустройств\w*|озеленени\w*)",
            focus_text,
            re.IGNORECASE
        )
    ):

        profile["work_types"].append(
            "ландшафт"
        )


    if (
        "визуализация"
        not in profile["work_types"]
        and re.search(
            r"(?:разработ\w*|подготов\w*|выполн\w*).{0,120}"
            r"визуализац\w*",
            focus_text,
            re.IGNORECASE
        )
    ):

        profile["work_types"].append(
            "визуализация"
        )


    # --------------------------------------------------------
    # Архитектурная графика
    #
    # Не считаем любой "чертёж" архитектурной графикой.
    # Нужен архитектурный / проектный контекст.
    # --------------------------------------------------------

    architectural_context = re.search(
        r"(архитектур\w*|"
        r"дизайн\w*|"
        r"эскиз\w*|"
        r"архитектурн\w+\s+проект\w*|"
        r"дизайн[-\s]?проект|"
        r"проектирован\w*)",
        title_lower,
        re.IGNORECASE
    )


    graphical_result = re.search(
        r"(?:разработ\w*|подготов\w*|выполн\w*).{0,120}"
        r"(?:графическ\w+\s+материал|"
        r"чертеж\w*|"
        r"эскиз\w*|"
        r"визуализац\w*)",
        focus_text,
        re.IGNORECASE
    )


    if (
        architectural_context
        and graphical_result
    ):

        profile["work_types"].append(
            "архитектурная графика"
        )


    # Убираем дубликаты
    profile["work_types"] = list(
        dict.fromkeys(
            profile["work_types"]
        )
    )


    # ========================================================
    # РЕЗУЛЬТАТЫ / МАТЕРИАЛЫ
    # ========================================================

    deliverable_patterns = {

        "чертежи": r"(?:разработ\w*|подготов\w*|выполн\w*|переда\w*|представ\w*).{0,120}чертеж\w*",

        "графические материалы": r"(?:разработ\w*|подготов\w*|сформир\w*|переда\w*|представ\w*).{0,120}графическ\w+\s+материал",

        "эскизный проект": r"(?:разработ\w*|подготов\w*|выполн\w*).{0,120}эскизн\w*\s+проект\w*",

        "визуализации": r"(?:разработ\w*|подготов\w*|выполн\w*|переда\w*).{0,120}визуализац\w*",

        "проект планировки": r"проект\w*\s+планировк\w*",

        "проект межевания": r"проект\w*.{0,80}межевани\w*",

        "генеральный план": r"генеральн\w*\s+план\w*",

        "схема": r"схем\w*.{0,80}(?:планировочн\w*|организаци\w*)",

                "правила землепользования и застройки":
            r"правил\w*\s+землепользован\w*\s+и\s+застройк\w*",

        "кадастровые материалы":
            r"(?:кадастров\w*|межев\w*).{0,120}"
            r"(?:план\w*|материал\w*|документ\w*)",

        "сведения в ЕГРН":
            r"(?:внесени\w*|внесен\w*).{0,120}"
            r"егрн",
    }


    for name, pattern in deliverable_patterns.items():

        if re.search(
            pattern,
            focus_text,
            re.IGNORECASE
        ):

            profile["deliverables"].append(
                name
            )


    # ========================================================
    # СРОК ВЫПОЛНЕНИЯ
    # ========================================================

    completion_patterns = [

        r"(?:срок\w*|срок\s+исполнени\w*).{0,500}"
        r"(?:до\s+)?"
        r"\d{1,2}[./]\d{1,2}[./]\d{4}",

        r"(?:выполн\w*|оказан\w*).{0,300}"
        r"(?:до\s+)?"
        r"\d{1,2}[./]\d{1,2}[./]\d{4}",

        r"(?:\d{1,2}[./]\d{1,2}[./]\d{4}).{0,250}"
        r"(?:срок\w*|выполн\w*|оказан\w*)",

        r"до\s+\d{1,2}\s+"
        r"(?:января|февраля|марта|апреля|мая|июня|"
        r"июля|августа|сентября|октября|ноября|декабря)"
        r"\s+\d{4}",
    ]


    profile["completion_terms"] = (
        extract_evidence_snippets(
            contract_text,
            completion_patterns,
            max_snippets=5
        )
    )


    # ========================================================
    # ДОРАБОТКИ
    # ========================================================

    unlimited_revision_patterns = [

        r"без\s+ограничени\w*.{0,100}"
        r"(?:количеств\w*\s+)?"
        r"(?:доработ\w*|изменени\w*|замечан\w*)",

        r"неограниченн\w*\s+количеств\w*"
        r".{0,100}"
        r"(?:доработ\w*|изменени\w*|замечан\w*)",

        r"количеств\w*\s+доработ\w*"
        r".{0,100}"
        r"не\s+огранич\w*",

        r"до\s+полного\s+согласован\w*",
    ]


    free_revision_patterns = [

        r"без\s+дополнительн\w*\s+оплат\w*"
        r".{0,160}"
        r"(?:вносить|внести|доработ\w*|измен\w*|корректир\w*)",

        r"(?:вносить|внести|доработ\w*|измен\w*|корректир\w*)"
        r".{0,160}"
        r"без\s+дополнительн\w*\s+оплат\w*",
    ]


    profile["revision_terms"] = (
        extract_evidence_snippets(
            contract_text,
            unlimited_revision_patterns,
            max_snippets=5
        )
    )

    free_revision_terms = (
        extract_evidence_snippets(
            contract_text,
            free_revision_patterns,
            max_snippets=5
        )
    )

    if free_revision_terms:

        profile["revision_terms"].extend(
            free_revision_terms
        )


    # Дубликаты
    profile["revision_terms"] = list(
        dict.fromkeys(
            profile["revision_terms"]
        )
    )


    profile["risks"]["unlimited_revisions"] = bool(
        extract_evidence_snippets(
            contract_text,
            unlimited_revision_patterns,
            max_snippets=1
        )
    )

    profile["risks"]["free_revisions"] = bool(
        free_revision_terms
    )


    # ========================================================
    # ВНЕШНИЕ СОГЛАСОВАНИЯ
    # ========================================================

    approval_patterns = [

        r"(?:согласоват\w*|согласован\w*).{0,180}"
        r"(?:администраци\w*|"
        r"министерств\w*|"
        r"ведомств\w*|"
        r"орган\w*\s+местного\s+самоуправлен\w*|"
        r"сетев\w*\s+организаци\w*|"
        r"ресурсоснабжающ\w*\s+организаци\w*)",

        r"(?:получит\w*|получен\w*|обеспеч\w*).{0,150}"
        r"(?:согласован\w*|разрешен\w*).{0,150}"
        r"(?:администраци\w*|"
        r"министерств\w*|"
        r"ведомств\w*|"
        r"сетев\w*|"
        r"ресурсоснабжающ\w*)",
    ]


    profile["approval_terms"] = (
        extract_evidence_snippets(
            contract_text + "\n" + work_text,
            approval_patterns,
            max_snippets=5
        )
    )


    profile["risks"]["approvals"] = bool(
        profile["approval_terms"]
    )


    return profile

# ============================================================
# ПОЛУЧЕНИЕ УЗЛА ОКПД2
# ============================================================

def get_node_for_code(page, code):

    titles = page.locator(".dynatree-title")

    for i in range(titles.count()):

        element = titles.nth(i)

        try:

            if not element.is_visible():
                continue

            text = element.inner_text().strip()

            if ":" not in text:
                continue

            actual_code = text.split(
                ":",
                1
            )[0].strip()

            if actual_code == code:

                return element.locator(
                    "xpath=ancestor::span[contains(@class,'dynatree-node')][1]"
                )

        except Exception:
            continue

    return None

def extract_completion_terms(text):
    patterns = [
        r"срок\s+выполнения\s+работ[^.]{0,250}",
        r"срок\s+оказания\s+услуг[^.]{0,250}",
        r"срок\s+выполнения[^.]{0,250}",
        r"дата\s+окончания\s+оказания\s+услуг[^.]{0,250}",
        r"окончани\w*\s+срока\s+выполнения[^.]{0,250}",
    ]

    results = []

    for pattern in patterns:
        matches = re.findall(
            pattern,
            text,
            re.IGNORECASE
        )

        for match in matches:
            clean = re.sub(r"\s+", " ", match).strip()

            if clean and clean not in results:
                results.append(clean)

    return results[:5]

# ============================================================
# ПОЛУЧЕНИЕ НОМЕРА СТРАНИЦЫ
# ============================================================

def get_page_number(page):

    try:

        parsed = urlparse(
            page.url
        )

        query = parse_qs(
            parsed.query
        )

        value = query.get(
            "pageNumber",
            ["1"]
        )[0]

        return int(value)

    except Exception:

        return 1


# ============================================================
# ПЕРЕХОД НА СТРАНИЦУ РЕЗУЛЬТАТОВ
# ============================================================

def go_to_results_page(
    page,
    page_number
):

    parsed = urlparse(
        page.url
    )

    query = parse_qs(
        parsed.query,
        keep_blank_values=True
    )

    query["pageNumber"] = [
        str(page_number)
    ]

    new_query = urlencode(
        query,
        doseq=True
    )

    new_url = urlunparse(
        (
            parsed.scheme,
            parsed.netloc,
            parsed.path,
            parsed.params,
            new_query,
            parsed.fragment
        )
    )

    print(
        f"\nПереходим напрямую "
        f"на страницу #{page_number}"
    )

    print(new_url)

    try:

        page.goto(
            new_url,
            wait_until="domcontentloaded",
            timeout=120000
        )

        page.wait_for_timeout(
            5000
        )

        current_page = get_page_number(
            page
        )

        if current_page != page_number:

            print(
                f"⚠ Ожидалась страница "
                f"#{page_number}, "
                f"но ЕИС открыла "
                f"страницу #{current_page}"
            )

            return False

        print(
            f"✓ Открыта страница "
            f"#{page_number}"
        )

        return True

    except Exception as e:

        print(
            f"✗ Ошибка перехода "
            f"на страницу #{page_number}: "
            f"{e}"
        )

        return False


# ============================================================
# ПОЛУЧЕНИЕ ССЫЛОК НА ЗАКУПКИ
# ============================================================

def get_tender_link_info(
    href
):

    if not href:
        return None


    # --------------------------------------------------------
    # 223-ФЗ
    # --------------------------------------------------------

    if "purchaseNoticeNumber" in href:

        match = re.search(
            r"purchaseNoticeNumber=([^&]+)",
            href
        )

        if match:

            number = match.group(1)

            return {
                "number": number,
                "law": "223-ФЗ",
                "url": href
            }


    # --------------------------------------------------------
    # 44-ФЗ
    # --------------------------------------------------------

    if "regNumber=" in href:

        match = re.search(
            r"regNumber=([^&]+)",
            href
        )

        if match:

            number = match.group(1)

            return {
                "number": number,
                "law": "44-ФЗ",
                "url": href
            }


    return None


# ============================================================
# СБОР ТЕНДЕРОВ С ТЕКУЩЕЙ СТРАНИЦЫ
# ============================================================

def collect_tenders_from_current_page(
    page
):

    all_links = page.locator(
        "a"
    )

    print(
        f"\nСсылок на текущей странице: "
        f"{all_links.count()}"
    )

    tenders = {}


    for i in range(
        all_links.count()
    ):

        link = all_links.nth(i)

        try:

            if not link.is_visible():
                continue

            href = link.get_attribute(
                "href"
            )

            if not href:
                continue

            tender_info = (
                get_tender_link_info(
                    href
                )
            )

            if not tender_info:
                continue

            number = tender_info[
                "number"
            ]

            if number not in tenders:

                tenders[
                    number
                ] = tender_info


        except Exception:
            continue


    return list(
        tenders.values()
    )


# ============================================================
# ПОИСК КОЛИЧЕСТВА ЗАПИСЕЙ
# ============================================================

def get_records_count(
    page
):

    text = page.locator(
        "body"
    ).inner_text()

    patterns = [
        r"Количество найденных записей:\s*(\d+)",
        r"Найдено записей:\s*(\d+)",
        r"Найдено:\s*(\d+)",
        r"Всего найдено:\s*(\d+)",
    ]


    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            re.IGNORECASE
        )

        if match:

            try:

                return int(
                    match.group(1)
                )

            except Exception:
                pass


    return 0


# ============================================================
# НАЧАЛО ИЗМЕНЁННОГО БЛОКА
#
# ПОЛУЧЕНИЕ СОДЕРЖИМОГО ФАЙЛА БЕЗ СОХРАНЕНИЯ НА ДИСК
# ============================================================

def get_file_response(
    page,
    file_url
):

    print(
        "\nПолучаем содержимое:"
    )

    print(
        file_url
    )


    try:

        response = page.request.get(
            file_url,
            timeout=120000
        )


        print(
            f"HTTP статус: "
            f"{response.status}"
        )


        content_type = (
            response.headers.get(
                "content-type",
                ""
            )
        )


        print(
            "Content-Type:",
            content_type
        )


        print(
            "Content-Length:",
            response.headers.get(
                "content-length",
                "не указан"
            )
        )


        if not response.ok:

            print(
                f"❌ Сервер вернул ошибку: "
                f"{response.status}"
            )

            return None


        body = response.body()


        print(
            f"✓ Получено байт: "
            f"{len(body)}"
        )


        return {
            "status": response.status,
            "headers": response.headers,
            "body": body
        }


    except Exception as e:

        print(
            f"❌ Ошибка получения файла: "
            f"{e}"
        )

        return None


# ============================================================
# ОПРЕДЕЛЯЕМ ФАКТИЧЕСКИЙ ФОРМАТ ФАЙЛА
#
# Мы не доверяем Content-Type, потому что ЕИС у нас отдаёт
# application/download.
# ============================================================

# ============================================================
# ОПРЕДЕЛЯЕМ ФАКТИЧЕСКИЙ ФОРМАТ ФАЙЛА
#
# Проверяем содержимое файла, а не только его название.
# Это особенно важно для ЕИС, потому что некоторые старые
# документы приходят без расширения и определяются как OLE.
# ============================================================

def detect_file_format(
    body,
    file_name=""
):

    # --------------------------------------------------------
    # PDF
    # --------------------------------------------------------

    if body.startswith(
        b"%PDF"
    ):

        return "pdf"


    # --------------------------------------------------------
    # ZIP / DOCX / XLSX
    #
    # DOCX и XLSX являются ZIP-контейнерами.
    # --------------------------------------------------------

    if body.startswith(
        b"PK"
    ):

        try:

            import zipfile

            with zipfile.ZipFile(
                BytesIO(body),
                "r"
            ) as archive:

                names = archive.namelist()


                # ------------------------------------------------
                # DOCX
                # ------------------------------------------------

                if (
                    "[Content_Types].xml"
                    in names
                    and
                    any(
                        name.startswith(
                            "word/"
                        )
                        for name in names
                    )
                ):

                    return "docx"


                # ------------------------------------------------
                # XLSX
                # ------------------------------------------------

                if (
                    "[Content_Types].xml"
                    in names
                    and
                    any(
                        name.startswith(
                            "xl/"
                        )
                        for name in names
                    )
                ):

                    return "xlsx"


                return "zip"


        except Exception:

            pass


    # --------------------------------------------------------
    # Старый Microsoft Office:
    #
    # DOC / XLS = OLE Compound File
    #
    # Иногда ЕИС не указывает расширение вообще.
    # Поэтому определяем тип по внутренним потокам OLE.
    # --------------------------------------------------------

    if body.startswith(
        bytes.fromhex(
            "D0CF11E0A1B11AE1"
        )
    ):

        name_lower = (
            file_name.lower()
        )


        # ----------------------------------------------------
        # Сначала используем расширение, если оно есть.
        # ----------------------------------------------------

        if name_lower.endswith(
            ".xls"
        ):

            return "xls"


        if name_lower.endswith(
            ".doc"
        ):

            return "doc"


        # ----------------------------------------------------
        # Если расширения нет — исследуем OLE.
        # ----------------------------------------------------

        try:

            ole = olefile.OleFileIO(
                BytesIO(body)
            )


            streams = ole.listdir(
                streams=True,
                storages=False
            )


            stream_names = [
                item[-1].lower()
                for item in streams
                if item
            ]


            # ------------------------------------------------
            # Старый Word
            #
            # Главный поток Word-документа:
            # WordDocument
            # ------------------------------------------------

            if "worddocument" in stream_names:

                ole.close()

                return "doc"


            # ------------------------------------------------
            # Старый Excel
            #
            # Основной поток обычно называется:
            # Workbook
            #
            # В старых версиях также встречается:
            # Book
            # ------------------------------------------------

            if (
                "workbook" in stream_names
                or
                "book" in stream_names
            ):

                ole.close()

                return "xls"


            ole.close()


        except Exception as e:

            print(
                "⚠ Не удалось определить "
                f"тип OLE-файла: {e}"
            )


        # ----------------------------------------------------
        # OLE-файл есть, но это не Word и не Excel.
        # ----------------------------------------------------

        return "ole"


    # --------------------------------------------------------
    # Текстовые файлы
    # --------------------------------------------------------

    try:

        body.decode(
            "utf-8"
        )

        return "txt"

    except Exception:

        pass


    # --------------------------------------------------------
    # Неизвестный формат
    # --------------------------------------------------------

    return "unknown"


# ============================================================
# ИЗВЛЕКАЕМ ТЕКСТ ИЗ PDF
# ============================================================

def extract_text_from_pdf(
    body
):

    try:

        reader = PdfReader(
            BytesIO(body)
        )


        pages_text = []


        print(
            f"Количество страниц PDF: "
            f"{len(reader.pages)}"
        )


        for page_number, pdf_page in enumerate(
            reader.pages,
            start=1
        ):

            try:

                text = (
                    pdf_page.extract_text()
                    or ""
                )


                pages_text.append(
                    text
                )


            except Exception as e:

                print(
                    f"⚠ Ошибка страницы "
                    f"PDF #{page_number}: "
                    f"{e}"
                )


        return "\n".join(
            pages_text
        ).strip()


    except Exception as e:

        print(
            f"❌ Ошибка чтения PDF: "
            f"{e}"
        )

        return ""

# ============================================================
# ИЗВЛЕКАЕМ ТЕКСТ ИЗ СТАРОГО DOC
# ============================================================

def extract_text_from_doc(
    body
):

    try:

        result = extract_legacy_doc_text(
            body
        )

        text = (
            result.text
            if hasattr(result, "text")
            else str(result)
        )

        return text.strip()


    except Exception as e:

        print(
            f"❌ Ошибка чтения DOC: "
            f"{e}"
        )

        return ""

# ============================================================
# ИЗВЛЕКАЕМ ТЕКСТ ИЗ DOCX
# ============================================================

def extract_text_from_docx(
    body
):

    try:

        document = Document(
            BytesIO(body)
        )


        parts = []


        # ----------------------------------------------------
        # Обычные абзацы
        # ----------------------------------------------------

        for paragraph in document.paragraphs:

            text = (
                paragraph.text
                .strip()
            )


            if text:

                parts.append(
                    text
                )


        # ----------------------------------------------------
        # Таблицы
        # ----------------------------------------------------

        for table_number, table in enumerate(
            document.tables,
            start=1
        ):

            parts.append(
                f"\n[Таблица {table_number}]"
            )


            for row in table.rows:

                cells = []


                for cell in row.cells:

                    cells.append(
                        cell.text.strip()
                    )


                row_text = " | ".join(
                    cells
                )


                if row_text.strip():

                    parts.append(
                        row_text
                    )


        return "\n".join(
            parts
        ).strip()


    except Exception as e:

        print(
            f"❌ Ошибка чтения DOCX: "
            f"{e}"
        )

        return ""


# ============================================================
# ИЗВЛЕКАЕМ ТЕКСТ ИЗ XLSX
# ============================================================

def extract_text_from_xlsx(
    body
):

    try:

        workbook = load_workbook(
            BytesIO(body),
            read_only=True,
            data_only=True
        )


        parts = []


        for sheet in workbook.worksheets:

            print(
                f"Лист Excel: "
                f"{sheet.title}"
            )


            parts.append(
                f"\n[ЛИСТ: {sheet.title}]"
            )


            for row in sheet.iter_rows(
                values_only=True
            ):

                values = []


                for value in row:

                    if value is None:

                        values.append("")

                    else:

                        values.append(
                            str(value)
                        )


                # Убираем полностью пустые строки

                if not any(
                    value.strip()
                    for value in values
                ):

                    continue


                row_text = " | ".join(
                    values
                )


                parts.append(
                    row_text
                )


        workbook.close()


        return "\n".join(
            parts
        ).strip()


    except Exception as e:

        print(
            f"❌ Ошибка чтения XLSX: "
            f"{e}"
        )

        return ""

# ============================================================
# ИЗВЛЕКАЕМ ТЕКСТ ИЗ СТАРОГО XLS
# ============================================================

def extract_text_from_xls(
    body
):

    try:

        workbook = xlrd.open_workbook(
            file_contents=body,
            on_demand=True
        )

        parts = []


        for sheet in workbook.sheets():

            print(
                f"Лист Excel: "
                f"{sheet.name}"
            )

            parts.append(
                f"\n[ЛИСТ: {sheet.name}]"
            )


            for row_index in range(
                sheet.nrows
            ):

                values = []


                for col_index in range(
                    sheet.ncols
                ):

                    cell = sheet.cell(
                        row_index,
                        col_index
                    )

                    value = cell.value


                    if value is None:

                        values.append("")

                    else:

                        values.append(
                            str(value)
                        )


                # Убираем полностью пустые строки

                if not any(
                    value.strip()
                    for value in values
                ):

                    continue


                row_text = " | ".join(
                    values
                )


                parts.append(
                    row_text
                )


        workbook.release_resources()


        return "\n".join(
            parts
        ).strip()


    except Exception as e:

        print(
            f"❌ Ошибка чтения XLS: "
            f"{e}"
        )

        return ""

    # ============================================================
# ИЗВЛЕКАЕМ ТЕКСТ ИЗ ZIP-АРХИВА
#
# Некоторые файлы ЕИС приходят как ZIP.
# Внутри архива могут находиться DOCX, XLSX, DOC, XLS,
# PDF, TXT и другие документы.
#
# Ничего на диск не сохраняем.
# ============================================================

def extract_text_from_zip(
    body
):

    try:

        import zipfile


        with zipfile.ZipFile(
            BytesIO(body),
            "r"
        ) as archive:

            members = archive.infolist()


            print(
                f"Файлов внутри ZIP: "
                f"{len(members)}"
            )


            parts = []


            for member in members:

                # ------------------------------------------------
                # Пропускаем папки
                # ------------------------------------------------

                if member.is_dir():

                    continue


                file_name = member.filename


                print()
                print(
                    f"  ZIP-файл: "
                    f"{file_name}"
                )


                try:

                    nested_body = archive.read(
                        member
                    )

                except Exception as e:

                    print(
                        f"  ❌ Не удалось прочитать "
                        f"{file_name}: {e}"
                    )

                    continue


                print(
                    f"  Размер: "
                    f"{len(nested_body)} байт"
                )


                # ------------------------------------------------
                # Определяем формат внутреннего файла
                # ------------------------------------------------

                nested_format = detect_file_format(
                    nested_body,
                    file_name
                )


                print(
                    f"  Формат внутри ZIP: "
                    f"{nested_format}"
                )


                # ------------------------------------------------
                # Сам ZIP внутри ZIP пока не обрабатываем
                # бесконечно
                # ------------------------------------------------

                if nested_format == "zip":

                    print(
                        "  ⚠ Внутри найден ещё один ZIP."
                    )

                    continue


                # ------------------------------------------------
                # Извлекаем текст
                # ------------------------------------------------

                try:

                    result = extract_text_from_file(
                        nested_body,
                        file_name
                    )

                    nested_text = (
                        result.get(
                            "text",
                            ""
                        )
                        or ""
                    )


                except Exception as e:

                    print(
                        f"  ❌ Ошибка извлечения "
                        f"текста: {e}"
                    )

                    nested_text = ""


                print(
                    f"  Размер текста: "
                    f"{len(nested_text)}"
                )


                # ------------------------------------------------
                # Сохраняем только то, где есть текст
                # ------------------------------------------------

                if nested_text.strip():

                    parts.append(
                        "\n".join(
                            [
                                "----------------------------------------",
                                f"ФАЙЛ ВНУТРИ ZIP: {file_name}",
                                f"ФОРМАТ: {nested_format}",
                                "----------------------------------------",
                                nested_text
                            ]
                        )
                    )


            # ----------------------------------------------------
            # Объединяем текст всех внутренних документов
            # ----------------------------------------------------

            return "\n\n".join(
                parts
            ).strip()


    except Exception as e:

        print(
            f"❌ Ошибка чтения ZIP: "
            f"{e}"
        )

        return ""

# ============================================================
# ИЗВЛЕКАЕМ ТЕКСТ ИЗ TXT
# ============================================================

def extract_text_from_txt(
    body
):

    encodings = [
        "utf-8",
        "cp1251",
        "utf-16",
    ]


    for encoding in encodings:

        try:

            return body.decode(
                encoding
            ).strip()

        except Exception:

            continue


    return ""


# ============================================================
# ОБЩИЙ МЕТОД ИЗВЛЕЧЕНИЯ ТЕКСТА
# ============================================================

def extract_text_from_file(
    body,
    file_name
):

    file_format = detect_file_format(
        body,
        file_name
    )


    print(
        f"Определённый формат: "
        f"{file_format}"
    )


    # --------------------------------------------------------
    # PDF
    # --------------------------------------------------------

    if file_format == "pdf":

        text = extract_text_from_pdf(
            body
        )

        return {
            "format": "pdf",
            "text": text
        }


    # --------------------------------------------------------
    # DOCX
    # --------------------------------------------------------

    if file_format == "docx":

        text = extract_text_from_docx(
            body
        )

        return {
            "format": "docx",
            "text": text
        }


    # --------------------------------------------------------
    # XLSX
    # --------------------------------------------------------

    if file_format == "xlsx":

        text = extract_text_from_xlsx(
            body
        )

        return {
            "format": "xlsx",
            "text": text
        }


    # --------------------------------------------------------
    # TXT
    # --------------------------------------------------------

    if file_format == "txt":

        text = extract_text_from_txt(
            body
        )

        return {
            "format": "txt",
            "text": text
        }

    # --------------------------------------------------------
    # ZIP
    # --------------------------------------------------------

    if file_format == "zip":

        text = extract_text_from_zip(
            body
        )

        return {
            "format": "zip",
            "text": text
        }


    # --------------------------------------------------------
    # Старые DOC/XLS
    # --------------------------------------------------------

    # --------------------------------------------------------
# Старый DOC
# --------------------------------------------------------

    if file_format == "doc":

        text = extract_text_from_doc(
            body
        )

        return {
            "format": "doc",
            "text": text
        }


    # --------------------------------------------------------
    # Старый XLS
    # --------------------------------------------------------

    if file_format == "xls":

        text = extract_text_from_xls(
            body
        )

        return {
            "format": "xls",
            "text": text
        }


    # --------------------------------------------------------
    # Неизвестный OLE
    # --------------------------------------------------------

    if file_format == "ole":

        print(
            "⚠ Обнаружен OLE-файл, "
            "но расширение не определено."
        )

        return {
            "format": "ole",
            "text": ""
        }


    # --------------------------------------------------------
    # ZIP / неизвестный
    # --------------------------------------------------------

    print(
        "⚠ Не удалось определить формат "
        "документа."
    )


    return {
        "format": file_format,
        "text": ""
    }


# ============================================================
# КОНЕЦ ИЗМЕНЁННОГО БЛОКА
# ============================================================


# ============================================================
# ОСНОВНАЯ ПРОГРАММА
# ============================================================

with sync_playwright() as p:

    print(
        "Запускаем браузер..."
    )


    browser = p.chromium.launch(
        headless=False
    )


    context = browser.new_context(
        ignore_https_errors=True
    )


    page = context.new_page()


    # ========================================================
    # ОТКРЫВАЕМ РАСШИРЕННЫЙ ПОИСК
    # ========================================================

    print(
        "\nОткрываем расширенный поиск..."
    )


    page.goto(
        "https://zakupki.gov.ru/epz/order/extendedsearch/search.html",
        wait_until="domcontentloaded",
        timeout=120000
    )


    page.wait_for_timeout(
        5000
    )


    # ========================================================
    # ЦЕНА
    # ========================================================

    print("\n" + "=" * 60)
    print(
        "НАСТРАИВАЕМ ЦЕНУ"
    )
    print("=" * 60)


    price_title = page.get_by_text(
        "Цена",
        exact=True
    )


    print(
        f"Найдено элементов «Цена»: "
        f"{price_title.count()}"
    )


    for i in range(
        price_title.count()
    ):

        element = price_title.nth(i)

        try:

            if element.is_visible():

                print(
                    "Открываем блок «Цена»..."
                )


                element.click()


                page.wait_for_timeout(
                    500
                )


                break

        except Exception:
            continue


    price_input = page.locator(
        "#priceToGeneral"
    )


    print(
        f"Поле максимальной цены найдено: "
        f"{price_input.count()}"
    )


    price_input.fill(
        MAX_PRICE
    )


    print(
        f"Максимальная цена: "
        f"{MAX_PRICE}"
    )


    # ========================================================
    # ЭТАП ЗАКУПКИ
    # ========================================================

    print("\n" + "=" * 60)
    print(
        "НАСТРАИВАЕМ ЭТАП ЗАКУПКИ"
    )
    print("=" * 60)


    af = page.locator(
        "#af"
    )

    ca = page.locator(
        "#ca"
    )

    pc = page.locator(
        "#pc"
    )

    pa = page.locator(
        "#pa"
    )


    print(
        "\nТекущее состояние:"
    )


    print(
        "Подача заявок:",
        af.is_checked()
    )

    print(
        "Работа комиссии:",
        ca.is_checked()
    )

    print(
        "Закупка завершена:",
        pc.is_checked()
    )

    print(
        "Закупка отменена:",
        pa.is_checked()
    )


    if not af.is_checked():

        page.locator(
            'label[for="af"]'
        ).click()


    if ca.is_checked():

        page.locator(
            'label[for="ca"]'
        ).click()


    if pc.is_checked():

        page.locator(
            'label[for="pc"]'
        ).click()


    if pa.is_checked():

        page.locator(
            'label[for="pa"]'
        ).click()


    print(
        "\nИтог этапа:"
    )


    print(
        "Подача заявок:",
        af.is_checked()
    )

    print(
        "Работа комиссии:",
        ca.is_checked()
    )

    print(
        "Закупка завершена:",
        pc.is_checked()
    )

    print(
        "Закупка отменена:",
        pa.is_checked()
    )


    # ========================================================
    # ОКПД2
    # ========================================================

    print("\n" + "=" * 60)
    print(
        "НАСТРАИВАЕМ ОКПД2"
    )
    print("=" * 60)


    spravochniki = page.get_by_text(
        "Справочники",
        exact=True
    )


    print(
        f"Найдено «Справочники»: "
        f"{spravochniki.count()}"
    )


    for i in range(
        spravochniki.count()
    ):

        element = spravochniki.nth(i)

        try:

            if element.is_visible():

                element.click()

                page.wait_for_timeout(
                    700
                )

                break

        except Exception:
            continue


    okpd_anchor = page.locator(
        "#okpd2IdsAnchor"
    )


    print(
        f"ОКПД2 найден: "
        f"{okpd_anchor.count()}"
    )


    okpd_anchor.click()


    page.wait_for_timeout(
        1000
    )


    goods_search = page.locator(
        "#goodssearch"
    )


    for _ in range(20):

        if goods_search.is_visible():
            break

        page.wait_for_timeout(
            500
        )


    print(
        "Окно ОКПД2 открыто."
    )


    # ========================================================
    # ИЩЕМ 71.11
    # ========================================================

    print(
        "\nИщем код 71.11..."
    )


    goods_search.click()


    goods_search.fill(
        ""
    )


    goods_search.type(
        "71.11",
        delay=300
    )


    found_search_result = False


    for second in range(
        15
    ):

        page.wait_for_timeout(
            1000
        )


        titles = page.locator(
            ".dynatree-title"
        )


        for i in range(
            titles.count()
        ):

            element = titles.nth(i)


            try:

                if not element.is_visible():
                    continue


                text = (
                    element.inner_text()
                    .strip()
                )


                if text.startswith(
                    "71.11:"
                ):

                    found_search_result = True


                    print(
                        f"Код 71.11 появился "
                        f"через ~{second + 1} сек."
                    )


                    break


            except Exception:
                continue


        if found_search_result:
            break


    if not found_search_result:

        print(
            "❌ Код 71.11 не найден."
        )


        browser.close()


        raise SystemExit


    print(
        "Результаты поиска готовы."
    )


    # ========================================================
    # НАХОДИМ РОДИТЕЛЬСКИЙ КОД
    # ========================================================

    titles = page.locator(
        ".dynatree-title"
    )


    parent_title = None


    for i in range(
        titles.count()
    ):

        element = titles.nth(i)


        try:

            if not element.is_visible():
                continue


            text = (
                element.inner_text()
                .strip()
            )


            if text.startswith(
                "71.11:"
            ):

                parent_title = element


                print(
                    f"\nНайден родительский код: "
                    f"{text}"
                )


                break


        except Exception:
            continue


    if parent_title is None:

        print(
            "❌ Родительский код не найден."
        )


        browser.close()


        raise SystemExit


    parent_node = parent_title.locator(
        "xpath=ancestor::span[contains(@class,'dynatree-node')][1]"
    )


    parent_checkbox = parent_node.locator(
        ":scope > .dynatree-checkbox"
    )


    # ========================================================
    # ВЫБИРАЕМ 71.11
    # ========================================================

    node_class = (
        parent_node.get_attribute(
            "class"
        )
        or ""
    )


    if "dynatree-selected" not in node_class:

        parent_checkbox.click()

        page.wait_for_timeout(
            500
        )


    # ========================================================
    # РАСКРЫВАЕМ 71.11
    # ========================================================

    node_class = (
        parent_node.get_attribute(
            "class"
        )
        or ""
    )


    if "dynatree-expanded" not in node_class:

        expander = parent_node.locator(
            ":scope > .dynatree-expander"
        )


        if expander.count() > 0:

            expander.click()

            page.wait_for_timeout(
                1500
            )


    page.wait_for_timeout(
        1500
    )


    # ========================================================
    # РАСКРЫВАЕМ ВСЕ ВЕТКИ 71.11
    # ========================================================

    print("\nРаскрываем все ветки ОКПД2 71.11...")

    for round_number in range(10):

        expanded_any = False

        titles = page.locator(
            ".dynatree-title"
        )

        total_titles = titles.count()

        for i in range(total_titles):

            element = titles.nth(i)

            try:

                text = (
                    element.inner_text()
                    .strip()
                )

                if ":" not in text:
                    continue

                code = text.split(
                    ":",
                    1
                )[0].strip()

                # Нас интересует только семейство 71.11
                if not re.match(
                    r"^71\.11(?:\.|$)",
                    code
                ):
                    continue

                node = element.locator(
                    "xpath=ancestor::span[contains(@class,'dynatree-node')][1]"
                )

                if node.count() == 0:
                    continue

                node_class = (
                    node.get_attribute(
                        "class"
                    )
                    or ""
                )

                # Уже раскрыта
                if "dynatree-expanded" in node_class:
                    continue

                expander = node.locator(
                    ":scope > .dynatree-expander"
                )

                if expander.count() == 0:
                    continue

                try:

                    expander.click(
                        timeout=3000
                    )

                    page.wait_for_timeout(
                        300
                    )

                    expanded_any = True

                except Exception:
                    continue

            except Exception:

                continue

        if not expanded_any:
            break

        print(
            f"  Раунд раскрытия: "
            f"{round_number + 1}"
        )

        page.wait_for_timeout(
            500
        )


    page.wait_for_timeout(
        1000
    )


    # ========================================================
    # СОБИРАЕМ ВСЕ КОДЫ СЕМЕЙСТВА 71.11
    # ========================================================

    titles = page.locator(
        ".dynatree-title"
    )

    family_codes = []
    seen_codes = set()

    for i in range(
        titles.count()
    ):

        element = titles.nth(i)

        try:

            text = (
                element.inner_text()
                .strip()
            )

            if ":" not in text:
                continue

            code = text.split(
                ":",
                1
            )[0].strip()

            if not re.match(
                r"^71\.11(?:\.|$)",
                code
            ):
                continue

            if code in seen_codes:
                continue

            seen_codes.add(code)

            family_codes.append(
                (
                    code,
                    text
                )
            )

        except Exception:

            continue


    # Сортировка по числовым частям кода
    def okpd_sort_key(item):

        return tuple(
            int(part)
            for part in item[0].split(".")
        )


    family_codes.sort(
        key=okpd_sort_key
    )


    print(
        f"\nНайдено кодов семейства 71.11: "
        f"{len(family_codes)}"
    )

    for code, text in family_codes:

        print(
            f"   {text}"
        )


    # ========================================================
    # ПРОВЕРЯЕМ, НЕ ПРОПАЛИ ЛИ КОДЫ
    # ========================================================

    found_codes = {
        code
        for code, text in family_codes
    }


    missing_codes = [
        code
        for code in EXPECTED_OKPD2_CODES
        if code not in found_codes
    ]


    # Коды, которые пришлось искать отдельно
    # и которые мы уже выбрали во время отдельного поиска.
    recovered_selected_codes = set()


    if missing_codes:

        print(
            "\n⚠ В дереве ОКПД2 не найдены:"
        )

        for code in missing_codes:

            print(
                f"   {code}"
            )


        print(
            "\nПробуем найти отсутствующие коды "
            "через поиск ОКПД2..."
        )


        search_input = page.locator(
            "#goodssearch"
        )


        for missing_code in missing_codes:

            try:

                # ------------------------------------------------
                # Ищем конкретный отсутствующий код
                # ------------------------------------------------

                search_input.click()

                search_input.press(
                    "Control+A"
                )

                search_input.type(
                    missing_code,
                    delay=150
                )

                page.wait_for_timeout(
                    1200
                )


                # ------------------------------------------------
                # Получаем его узел
                # ------------------------------------------------

                node = get_node_for_code(
                    page,
                    missing_code
                )


                if node is None:

                    print(
                        f"  ❌ Не удалось найти: "
                        f"{missing_code}"
                    )

                    continue


                # ------------------------------------------------
                # Получаем название
                # ------------------------------------------------

                title_element = node.locator(
                    ":scope > .dynatree-title"
                )


                if title_element.count() > 0:

                    full_text = (
                        title_element.inner_text()
                        .strip()
                    )

                else:

                    full_text = missing_code


                print(
                    f"\n  Найден отсутствующий код:"
                    f"\n  {full_text}"
                )


                # ------------------------------------------------
                # СРАЗУ выбираем найденный код
                # ------------------------------------------------

                node_class = (
                    node.get_attribute(
                        "class"
                    )
                    or ""
                )


                if "dynatree-selected" in node_class:

                    print(
                        "  ✅ Код уже выбран."
                    )

                    recovered_selected_codes.add(
                        missing_code
                    )

                else:

                    checkbox = node.locator(
                        ":scope > .dynatree-checkbox"
                    )


                    if checkbox.count() == 0:

                        print(
                            "  ❌ Checkbox найден не был."
                        )

                        continue


                    print(
                        "  Выбираем checkbox..."
                    )


                    checkbox.click()


                    page.wait_for_timeout(
                        500
                    )


                    # ------------------------------------------------
                    # Повторно получаем node
                    # ------------------------------------------------

                    node = get_node_for_code(
                        page,
                        missing_code
                    )


                    if node is None:

                        print(
                            "  ❌ После клика узел "
                            "больше не найден."
                        )

                        continue


                    node_class = (
                        node.get_attribute(
                            "class"
                        )
                        or ""
                    )


                    # ------------------------------------------------
                    # Проверяем выбор
                    # ------------------------------------------------

                    if "dynatree-selected" in node_class:

                        print(
                            "  ✅ Код успешно выбран."
                        )

                        recovered_selected_codes.add(
                            missing_code
                        )

                    else:

                        print(
                            "  ❌ Код найден, "
                            "но выбрать его не удалось."
                        )

                        continue


                # ------------------------------------------------
                # Добавляем код в общий список
                # ------------------------------------------------

                family_codes.append(
                    (
                        missing_code,
                        full_text
                    )
                )


            except Exception as e:

                print(
                    f"  ❌ Ошибка поиска "
                    f"{missing_code}: {e}"
                )


        # ========================================================
        # ВОЗВРАЩАЕМ ОБЩИЙ ПОИСК 71.11
        # ========================================================

        try:

            search_input.click()

            search_input.press(
                "Control+A"
            )

            search_input.type(
                "71.11",
                delay=150
            )

            page.wait_for_timeout(
                1200
            )

        except Exception:

            pass


    # ========================================================
    # УБИРАЕМ ДУБЛИКАТЫ
    # ========================================================

    unique_family_codes = {}


    for code, text in family_codes:

        unique_family_codes[code] = text


    family_codes = list(
        unique_family_codes.items()
    )


    # ========================================================
    # СОРТИРУЕМ КОДЫ
    # ========================================================

    family_codes.sort(
        key=lambda item: tuple(
            int(part)
            for part in item[0].split(".")
        )
    )


    print(
        f"\nИтого кодов семейства 71.11 "
        f"после проверки: {len(family_codes)}"
    )


    # ========================================================
    # ПРОВЕРЯЕМ ПОЛНЫЙ СПИСОК
    # ========================================================

    if len(family_codes) != len(
        EXPECTED_OKPD2_CODES
    ):

        print(
            "❌ Ошибка: количество кодов "
            "отличается от ожидаемого."
        )

        print(
            "Ожидалось:",
            len(EXPECTED_OKPD2_CODES)
        )

        print(
            "Получено:",
            len(family_codes)
        )

        browser.close()
        exit()


    else:

        print(
            "✅ Все ожидаемые коды 71.11 найдены."
        )


    # ========================================================
    # ВЫБИРАЕМ КОДЫ
    # ========================================================

    selected_count = 0
    skipped_count = 0
    error_count = 0


    for code, text in family_codes:

        # Код был найден через отдельный поиск
        # и уже выбран там.
        if code in recovered_selected_codes:

            print(
                f"\nОбрабатываем: "
                f"{text}"
            )

            print(
                "  ✅ Уже выбран при восстановлении."
            )

            selected_count += 1

            continue

        if code in EXCLUDED_OKPD2:

            print(
                f"\nПропускаем исключение: "
                f"{text}"
            )

            skipped_count += 1

            continue


        print(
            f"\nОбрабатываем: "
            f"{text}"
        )


        node = get_node_for_code(
            page,
            code
        )


        if node is None:

            print(
                "  ❌ Узел не найден."
            )

            error_count += 1

            continue


        node_class = (
            node.get_attribute(
                "class"
            )
            or ""
        )


        if "dynatree-selected" in node_class:

            print(
                "  ✅ Уже выбран."
            )

            selected_count += 1

            continue


        checkbox = node.locator(
            ":scope > .dynatree-checkbox"
        )


        if checkbox.count() == 0:

            print(
                "  ❌ Checkbox не найден."
            )

            error_count += 1

            continue


        checkbox.click()


        page.wait_for_timeout(
            200
        )


        node = get_node_for_code(
            page,
            code
        )


        if node is None:

            error_count += 1

            continue


        node_class = (
            node.get_attribute(
                "class"
            )
            or ""
        )


        if "dynatree-selected" in node_class:

            print(
                "  ✅ УСПЕШНО ВЫБРАН."
            )

            selected_count += 1

        else:

            print(
                "  ❌ Не удалось выбрать."
            )

            error_count += 1


    # ========================================================
    # ПРОВЕРЯЕМ ИСКЛЮЧЕНИЯ
    # ========================================================

    for code in EXCLUDED_OKPD2:

        node = get_node_for_code(
            page,
            code
        )


        if node is None:
            continue


        node_class = (
            node.get_attribute(
                "class"
            )
            or ""
        )


        if "dynatree-selected" in node_class:

            checkbox = node.locator(
                ":scope > .dynatree-checkbox"
            )


            if checkbox.count() > 0:

                checkbox.click()

                page.wait_for_timeout(
                    200
                )


    # ========================================================
    # ФИНАЛЬНАЯ ПРОВЕРКА КОЛИЧЕСТВА ВЫБРАННЫХ КОДОВ
    # ========================================================

    expected_selected_count = len(
        [
            code
            for code in EXPECTED_OKPD2_CODES
            if code not in EXCLUDED_OKPD2
        ]
    )


    print()
    print("=" * 60)
    print("ПРОВЕРЯЕМ КОЛИЧЕСТВО ВЫБРАННЫХ КОДОВ")
    print("=" * 60)


    print(
        "Ожидается выбрать:",
        expected_selected_count
    )

    print(
        "Фактически выбрано:",
        selected_count
    )


    if selected_count != expected_selected_count:

        print()
        print(
            "❌ ОШИБКА: выбраны не все нужные "
            "коды ОКПД2."
        )

        print(
            "Разница:",
            expected_selected_count - selected_count
        )

        browser.close()
        exit()


    print(
        "✅ Все необходимые коды выбраны."
    )

    # ========================================================
    # КНОПКА ВЫБРАТЬ
    # ========================================================

    choose_buttons = page.get_by_text(
        "ВЫБРАТЬ",
        exact=True
    )


    for i in range(
        choose_buttons.count()
    ):

        button = choose_buttons.nth(i)


        try:

            if button.is_visible():

                button.click()

                page.wait_for_timeout(
                    1000
                )

                break

        except Exception:
            continue


    # ========================================================
    # ПРОВЕРЯЕМ ОКПД2
    # ========================================================

    # ========================================================
    # ПРОВЕРЯЕМ ОКПД2 ПОСЛЕ НАЖАТИЯ «ВЫБРАТЬ»
    # ========================================================

    okpd_codes_input = page.locator(
        "#okpd2IdsCodes"
    )


    if okpd_codes_input.count() == 0:

        print(
            "\n❌ ОШИБКА: #okpd2IdsCodes не найден."
        )

        browser.close()
        exit()


    selected_okpd_value = (
        okpd_codes_input.input_value()
    )


    print(
        "\nВыбранный код:",
        selected_okpd_value
    )


    actual_selected_codes = {
        code.strip()
        for code in selected_okpd_value.split(",")
        if code.strip()
    }


    expected_selected_codes = {
        code
        for code in EXPECTED_OKPD2_CODES
        if code not in EXCLUDED_OKPD2
    }


    missing_final_codes = (
        expected_selected_codes
        - actual_selected_codes
    )


    unexpected_final_codes = (
        actual_selected_codes
        - expected_selected_codes
    )


    print()
    print(
        "Фактически кодов в фильтре:",
        len(actual_selected_codes)
    )


    if missing_final_codes:

        print()
        print(
            "❌ В ФИЛЬТРЕ НЕ ХВАТАЕТ:"
        )

        for code in sorted(
            missing_final_codes,
            key=lambda value: tuple(
                int(part)
                for part in value.split(".")
            )
        ):

            print(
                "   ",
                code
            )


        browser.close()
        exit()


    if unexpected_final_codes:

        print()
        print(
            "⚠ В ФИЛЬТР ПОПАЛИ НЕОЖИДАННЫЕ КОДЫ:"
        )

        for code in sorted(
            unexpected_final_codes
        ):

            print(
                "   ",
                code
            )


        browser.close()
        exit()


    print()
    print(
        "✅ ФИЛЬТР ОКПД2 ПРОВЕРЕН."
    )

    print(
        "✅ Выбрано ровно:",
        len(actual_selected_codes),
        "кодов."
    )


    # ========================================================
    # ПРИМЕНЯЕМ ФИЛЬТРЫ
    # ========================================================

    print("\n" + "=" * 60)
    print(
        "ПРИМЕНЯЕМ ФИЛЬТРЫ"
    )
    print("=" * 60)


    apply_buttons = page.get_by_text(
        "ПРИМЕНИТЬ",
        exact=True
    )


    for i in range(
        apply_buttons.count()
    ):

        button = apply_buttons.nth(i)


        try:

            if button.is_visible():

                button.click()

                break

        except Exception:
            continue


    page.wait_for_timeout(
        7000
    )


    # ========================================================
    # ПОЛУЧАЕМ РЕЗУЛЬТАТЫ
    # ========================================================

    records_count = get_records_count(
        page
    )


    print(
        "\nКоличество найденных записей:",
        records_count
    )


    # ========================================================
    # СБОР ВСЕХ СТРАНИЦ
    # ========================================================

    print("\n" + "=" * 60)
    print(
        "НАЧИНАЕМ СБОР ВСЕХ СТРАНИЦ"
    )
    print("=" * 60)


    all_tenders = {}


    page_number = get_page_number(
        page
    )


    max_pages = 100


    seen_page_signatures = set()


    while page_number <= max_pages:

        print(
            "\n" + "=" * 70
        )


        print(
            f"СБОР РЕЗУЛЬТАТОВ — "
            f"СТРАНИЦА #{page_number}"
        )


        print(
            "=" * 70
        )


        current_tenders = (
            collect_tenders_from_current_page(
                page
            )
        )


        print(
            f"На странице найдено тендеров: "
            f"{len(current_tenders)}"
        )


        page_signature = tuple(
            sorted(
                tender["number"]
                for tender in current_tenders
                if tender.get("number")
            )
        )


        if page_signature in seen_page_signatures:

            print(
                "⚠ Эта страница уже была "
                "обработана. Останавливаемся."
            )

            break


        seen_page_signatures.add(
            page_signature
        )


        new_count = 0


        for tender in current_tenders:

            number = tender.get(
                "number"
            )


            if not number:
                continue


            if number not in all_tenders:

                all_tenders[
                    number
                ] = tender


                new_count += 1


        print(
            f"Новых уникальных тендеров "
            f"добавлено: {new_count}"
        )


        print(
            f"Всего уникальных тендеров "
            f"собрано: "
            f"{len(all_tenders)}"
        )


        if (
            records_count > 0
            and
            len(all_tenders) >= records_count
        ):

            print(
                f"\n✓ Собраны все ожидаемые "
                f"записи: "
                f"{len(all_tenders)} "
                f"из {records_count}"
            )

            break


        if len(current_tenders) == 0:

            print(
                "\n⚠ На странице нет "
                "тендеров. Останавливаемся."
            )

            break


        next_page = (
            page_number + 1
        )


        moved = go_to_results_page(
            page,
            next_page
        )


        if not moved:

            print(
                f"\n⚠ Не удалось перейти "
                f"на страницу #{next_page}."
            )

            break


        page_number = next_page


    # ========================================================
    # СПИСОК НАЙДЕННЫХ ТЕНДЕРОВ
    # ========================================================

    print("\n" + "=" * 70)


    print(
        "ВСЕ НАЙДЕННЫЕ ЗАКУПКИ"
    )


    print(
        "=" * 70
    )

    tenders = list(all_tenders.values())


    for i, tender in enumerate(
        all_tenders.values(),
        start=1
    ):

        print(
            f"\n{i}. "
            f"{tender['number']} "
            f"({tender['law']})"
        )


        print(
            f"   URL: "
            f"{tender['url']}"
        )


    # =========================================================
    # 14. ОБРАБОТКА ВСЕХ НАЙДЕННЫХ ЗАКУПОК
    # =========================================================

    from urllib.parse import urljoin
    import json


    # ---------------------------------------------------------
    # ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ:
    # пытаемся найти нормальную карточку закупки
    # ---------------------------------------------------------

    def get_real_tender_url(
        tender_page,
        tender
    ):
        """
        Определяет полноценный URL карточки закупки.

        Для 44-ФЗ карточка строится напрямую.

        Для 223-ФЗ сразу открываем обычную
        карточку закупки.
        """

        number = tender["number"]
        law = tender["law"]

        # =====================================================
        # 44-ФЗ
        # =====================================================

        if law == "44-ФЗ":

            return (
                "https://zakupki.gov.ru"
                "/epz/order/notice/ok20/view/common-info.html"
                f"?regNumber={number}"
            )

        # =====================================================
        # 223-ФЗ
        # =====================================================

        if law == "223-ФЗ":

            number = tender["number"]

            # Для 223-ФЗ сразу открываем
            # обычную карточку закупки.

            url = (
                "https://zakupki.gov.ru"
                "/epz/order/notice/notice223/common-info.html"
                f"?regNumber={number}"
            )

            print(
                "Ссылка на карточку 223-ФЗ:",
                url
            )

            return url

        # =====================================================
        # Неизвестный закон
        # =====================================================

        return urljoin(
            "https://zakupki.gov.ru",
            tender.get(
                "url",
                ""
            )
        )


    # ---------------------------------------------------------
    # ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ:
    # ищем ссылку на вкладку "Документы"
    # ---------------------------------------------------------

    def find_documents_url(tender_page):

        # =====================================================
        # 1. Сначала пытаемся найти обычную ссылку
        # =====================================================

        links = tender_page.locator("a")

        for i in range(links.count()):

            link = links.nth(i)

            try:
                href = link.get_attribute("href")
            except Exception:
                href = None

            if not href:
                continue

            full_url = urljoin(
                tender_page.url,
                href
            )

            full_url_lower = full_url.lower()

            # -------------------------------------------------
            # Прямая страница документов 223-ФЗ
            # -------------------------------------------------

            if (
                "/notice223/" in full_url_lower
                and "documents.html" in full_url_lower
            ):
                return full_url

            # -------------------------------------------------
            # Обычная страница документов 44-ФЗ
            # -------------------------------------------------

            if "documents" in full_url_lower:

                try:
                    text = link.inner_text().strip().lower()
                except Exception:
                    text = ""

                if (
                    text == "документы"
                    or "документ" in text
                ):
                    return full_url


        # =====================================================
        # 2. Для 223-ФЗ ищем noticeInfoId внутри HTML
        # =====================================================

        if "notice223" in tender_page.url.lower():

            try:

                html = tender_page.content()

                # Уже готовая ссылка на documents.html
                match = re.search(
                    r"/notice223/documents\.html\?noticeInfoId=(\d+)",
                    html,
                    flags=re.IGNORECASE
                )

                if match:

                    notice_info_id = match.group(1)

                    return (
                        "https://zakupki.gov.ru"
                        "/epz/order/notice/notice223/documents.html"
                        f"?noticeInfoId={notice_info_id}"
                    )


                # Просто noticeInfoId внутри HTML
                matches = re.findall(
                    r"noticeInfoId(?:=|\"|\'|:|\s)+(\d+)",
                    html,
                    flags=re.IGNORECASE
                )

                if matches:

                    # Убираем дубликаты
                    unique_ids = []

                    for value in matches:

                        if value not in unique_ids:
                            unique_ids.append(value)


                    if unique_ids:

                        notice_info_id = unique_ids[0]

                        print(
                            "Найден noticeInfoId 223-ФЗ:",
                            notice_info_id
                        )

                        return (
                            "https://zakupki.gov.ru"
                            "/epz/order/notice/notice223/documents.html"
                            f"?noticeInfoId={notice_info_id}"
                        )

            except Exception as e:

                print(
                    "⚠ Ошибка поиска noticeInfoId:",
                    e
                )


        # =====================================================
        # 3. Старый поиск для остальных случаев
        # =====================================================

        for i in range(links.count()):

            link = links.nth(i)

            try:
                text = link.inner_text().strip().lower()
            except Exception:
                continue

            if "документ" not in text:
                continue

            try:
                href = link.get_attribute("href")
            except Exception:
                href = None

            if not href:
                continue

            full_url = urljoin(
                tender_page.url,
                href
            )

            if "documents" in full_url.lower():
                return full_url

        return None


    # ---------------------------------------------------------
    # ПЫТАЕМСЯ ИЗВЛЕЧЬ ЦЕНУ
    # ---------------------------------------------------------

    def extract_price(card_text):

        number_pattern = (
            r"(\d[\d\s\u00A0]*[,.]\d{2})"
        )

        patterns = [

            # 223-ФЗ
            rf"Начальная\s*(?:\(\s*максимальная\s*\)\s*)?"
            rf"цена(?:\s+договора)?"
            rf"[^0-9]{{0,150}}"
            rf"{number_pattern}",

            # Другие варианты
            rf"Начальная цена"
            rf"[^0-9]{{0,150}}"
            rf"{number_pattern}",

            rf"Цена договора"
            rf"[^0-9]{{0,150}}"
            rf"{number_pattern}",

            rf"Цена контракта"
            rf"[^0-9]{{0,150}}"
            rf"{number_pattern}",

            rf"НМЦК"
            rf"[^0-9]{{0,150}}"
            rf"{number_pattern}",
        ]

        for pattern in patterns:

            match = re.search(
                pattern,
                card_text,
                flags=re.IGNORECASE
            )

            if match:
                return match.group(1).strip()

        return None


    # ---------------------------------------------------------
    # ПЫТАЕМСЯ НАЙТИ СРОК ПОДАЧИ ЗАЯВОК
    # ---------------------------------------------------------

    # ---------------------------------------------------------
    # ПЫТАЕМСЯ НАЙТИ СРОК ПОДАЧИ ЗАЯВОК
    # ---------------------------------------------------------

    def extract_application_deadline(card_text):

        date_pattern = (
            r"(\d{2}\.\d{2}\.\d{4})"
        )

        # ---------------------------------------------------------
        # Основные варианты
        # ---------------------------------------------------------

        patterns = [

            # 44-ФЗ
            rf"Окончание\s+срока\s+подачи\s+заявок"
            rf"[^0-9]{{0,150}}"
            rf"{date_pattern}",

            rf"Дата\s+и\s+время\s+окончания\s+подачи\s+заявок"
            rf"[^0-9]{{0,150}}"
            rf"{date_pattern}",

            # 223-ФЗ
            rf"Окончание\s+подачи\s+заявок"
            rf"[^0-9]{{0,150}}"
            rf"{date_pattern}",

            rf"Дата\s+окончания\s+подачи\s+заявок"
            rf"[^0-9]{{0,150}}"
            rf"{date_pattern}",

            rf"Дата\s+окончания\s+приема\s+заявок"
            rf"[^0-9]{{0,150}}"
            rf"{date_pattern}",

            rf"Срока\s+подачи\s+заявок"
            rf"[^0-9]{{0,150}}"
            rf"{date_pattern}",

            # Иногда текст сформирован без слова "окончания"
            rf"подачи\s+заявок"
            rf"[^0-9]{{0,150}}"
            rf"{date_pattern}",
        ]


        for pattern in patterns:

            match = re.search(
                pattern,
                card_text,
                flags=re.IGNORECASE
            )

            if match:

                return match.group(1)


        # ---------------------------------------------------------
        # РЕЗЕРВНЫЙ ПОИСК
        #
        # Ищем все даты рядом со словами:
        # "подач", "прием", "заявок"
        #
        # Это нужно прежде всего для 223-ФЗ, где структура
        # текста карточки может отличаться.
        # ---------------------------------------------------------

        fallback_patterns = [

            rf"(?:подач\w*|прием\w*)"
            rf"[^.\n]{{0,200}}?"
            rf"{date_pattern}",

            rf"{date_pattern}"
            rf"[^.\n]{{0,200}}?"
            rf"(?:подач\w*|прием\w*)",
        ]


        for pattern in fallback_patterns:

            match = re.search(
                pattern,
                card_text,
                flags=re.IGNORECASE
            )

            if match:

                # В первом варианте дата находится во второй
                # группе из-за самой конструкции.

                groups = match.groups()

                for value in groups:

                    if value and re.fullmatch(
                        r"\d{2}\.\d{2}\.\d{4}",
                        value
                    ):

                        return value


        return None


    # ---------------------------------------------------------
    # ИЩЕМ КОДЫ ОКПД2 В КАРТОЧКЕ
    # ---------------------------------------------------------

    def extract_okpd2_codes(card_text):

        found = re.findall(
            r"\b71\.11(?:\.\d+){0,3}\b",
            card_text
        )

        unique_codes = []

        for code in found:

            if code not in unique_codes:
                unique_codes.append(code)

        return unique_codes


    # ---------------------------------------------------------
    # ПЫТАЕМСЯ НАЙТИ НАЗВАНИЕ ЗАКУПКИ
    # ---------------------------------------------------------

    def extract_tender_title(card_text):

        patterns = [
            r"Наименование объекта закупки\s*[:\-]?\s*(.+)",
            r"Наименование закупки\s*[:\-]?\s*(.+)",
            r"Объект закупки\s*[:\-]?\s*(.+)",
        ]

        for pattern in patterns:

            matches = re.findall(
                pattern,
                card_text,
                flags=re.IGNORECASE
            )

            for value in matches:

                value = value.strip()

                if len(value) < 10:
                    continue

                if len(value) > 500:
                    value = value[:500]

                return value

        return None


    # ---------------------------------------------------------
    # КЛИКАЕМ "ПОКАЗАТЬ БОЛЬШЕ", ЕСЛИ ЕСТЬ
    # ---------------------------------------------------------

    def expand_all_documents(tender_page):

        for _ in range(10):

            buttons = tender_page.get_by_text(
                "Показать больше",
                exact=True
            )

            clicked = False

            for i in range(buttons.count()):

                button = buttons.nth(i)

                try:

                    if not button.is_visible():
                        continue

                    button.click(
                        timeout=3000
                    )

                    tender_page.wait_for_timeout(
                        500
                    )

                    clicked = True

                    print(
                        "  Нажали «Показать больше»."
                    )

                    break

                except Exception:
                    continue

            if not clicked:
                break


    # ---------------------------------------------------------
    # ПОЛУЧАЕМ ВСЕ ПРИКРЕПЛЁННЫЕ ФАЙЛЫ
    # ---------------------------------------------------------

        # ---------------------------------------------------------
    # ПОЛУЧАЕМ ВСЕ ПРИКРЕПЛЁННЫЕ ФАЙЛЫ
    # ---------------------------------------------------------

    def get_document_files(tender_page):

        # =====================================================
        # Сначала раскрываем всё, что можно
        # =====================================================

        expand_all_documents(
            tender_page
        )


        files = []


        # =====================================================
        # 223-ФЗ
        #
        # На странице 223-ФЗ реальные документы находятся
        # непосредственно в ссылках:
        #
        # /223/filestore/public/1.0/download/fz223/file.html
        #
        # Поэтому .blockFilesTabDocs здесь не используется.
        # =====================================================

        if "notice223" in tender_page.url.lower():

            print()
            print(
                "Определяем файлы для 223-ФЗ..."
            )


            links = tender_page.locator(
                "a[href*='/223/filestore/']"
            )


            print(
                "Найдено ссылок на файлы 223-ФЗ:",
                links.count()
            )


            for i in range(
                links.count()
            ):

                link = links.nth(i)


                try:

                    href = link.get_attribute(
                        "href"
                    )

                except Exception:

                    href = None


                if not href:
                    continue


                full_url = urljoin(
                    tender_page.url,
                    href
                )


                # На всякий случай проверяем,
                # что это действительно файл ЕИС 223-ФЗ

                if "/223/filestore/" not in full_url:

                    continue


                try:

                    name = (
                        link.inner_text()
                        .strip()
                    )

                except Exception:

                    name = ""


                # -------------------------------------------------
                # Убираем ситуации, когда ссылка есть,
                # но название пустое
                # -------------------------------------------------

                if not name:

                    # Иногда название может лежать
                    # в title атрибуте

                    try:

                        name = (
                            link.get_attribute(
                                "title"
                            )
                            or ""
                        ).strip()

                    except Exception:

                        name = ""


                # -------------------------------------------------
                # Проверяем дубликаты
                # -------------------------------------------------

                exists = any(
                    item["url"] == full_url
                    for item in files
                )


                if exists:

                    continue


                files.append(
                    {
                        "name": (
                            name
                            or
                            "Без названия"
                        ),
                        "url": full_url
                    }
                )


            print(
                "Реальных файлов 223-ФЗ:",
                len(files)
            )


            for index, file_info in enumerate(
                files,
                start=1
            ):

                print(
                    f"  {index}. "
                    f"{file_info['name']}"
                )

                print(
                    f"     {file_info['url']}"
                )


            return files


        # =====================================================
        # 44-ФЗ
        #
        # Здесь оставляем старый проверенный механизм.
        # =====================================================

        body_text = tender_page.locator(
            "body"
        ).inner_text()


        if "Прикрепленные файлы" not in body_text:

            print(
                "  ⚠ Блок «Прикрепленные файлы» не найден."
            )


        containers = tender_page.locator(
            ".blockFilesTabDocs"
        )


        if containers.count() == 0:

            print(
                "  ⚠ .blockFilesTabDocs не найден."
            )

            return []


        container = containers.first


        links = container.locator(
            "a"
        )


        for i in range(
            links.count()
        ):

            link = links.nth(i)


            try:

                href = link.get_attribute(
                    "href"
                )

            except Exception:

                href = None


            if not href:

                continue


            full_url = urljoin(
                tender_page.url,
                href
            )


            # Нас интересуют реальные файлы ЕИС

            if (
                "/filestore/" not in full_url
                and
                "/download/" not in full_url
            ):

                continue


            try:

                name = (
                    link.inner_text()
                    .strip()
                )

            except Exception:

                name = ""


            # -------------------------------------------------
            # Убираем дубликаты
            # -------------------------------------------------

            exists = any(
                item["url"] == full_url
                for item in files
            )


            if exists:

                continue


            files.append(
                {
                    "name": (
                        name
                        or
                        "Без названия"
                    ),
                    "url": full_url
                }
            )


        return files


    # =========================================================
    # ОБЩИЙ СПИСОК ГОТОВЫХ ДАННЫХ
    # =========================================================

    processed_tenders = []


    print()
    print("=" * 70)
    print("НАЧИНАЕМ ОБРАБОТКУ ВСЕХ НАЙДЕННЫХ ЗАКУПОК")
    print("=" * 70)

    print()
    print(
        f"Всего закупок для обработки: {len(tenders)}"
    )


    # =========================================================
    # ОСНОВНОЙ ЦИКЛ ПО ВСЕМ ТЕНДЕРАМ
    # =========================================================

    for index, tender in enumerate(
        tenders,
        start=1
    ):

        print()
        print("=" * 70)
        print(
            f"ОБРАБОТКА ЗАКУПКИ {index}/{len(tenders)}"
        )
        print("=" * 70)

        print(
            "Номер:",
            tender["number"]
        )

        print(
            "Закон:",
            tender["law"]
        )

        tender_page = context.new_page()

        processed = {
            "number": tender["number"],
            "law": tender["law"],
            "notice_id": tender.get("notice_id"),
            "search_url": tender["url"],
            "tender_url": None,
            "documents_url": None,
            "title": None,
            "price": None,
            "application_deadline": None,
            "okpd2_codes": [],
            "card_text": "",
            "files": [],
            "documents_text": "",
            "total_files": 0,
            "loaded_files": 0,
            "status": "ok",
        }

        try:

            # -----------------------------------------------------
            # 1. НАХОДИМ НОРМАЛЬНУЮ КАРТОЧКУ
            # -----------------------------------------------------

            print()
            print(
                "Определяем нормальную ссылку на карточку..."
            )

            real_tender_url = get_real_tender_url(
                tender_page,
                tender
            )

            processed[
                "tender_url"
            ] = real_tender_url

            print(
                "Карточка:",
                real_tender_url
            )


            # -----------------------------------------------------
            # 2. ОТКРЫВАЕМ КАРТОЧКУ
            # -----------------------------------------------------

            tender_page.goto(
                real_tender_url,
                wait_until="domcontentloaded",
                timeout=120000
            )

            tender_page.wait_for_timeout(
                4000
            )

            print(
                "✓ Карточка открыта."
            )


            # -----------------------------------------------------
            # 3. ПОЛУЧАЕМ ТЕКСТ КАРТОЧКИ
            # -----------------------------------------------------

            card_text = tender_page.locator(
                "body"
            ).inner_text()

            processed[
                "card_text"
            ] = card_text

            print(
                "Размер текста карточки:",
                len(card_text)
            )


            # -----------------------------------------------------
            # 4. ИЗВЛЕКАЕМ ОСНОВНЫЕ ДАННЫЕ
            # -----------------------------------------------------

            processed[
                "title"
            ] = extract_tender_title(
                card_text
            )

            processed[
                "price"
            ] = extract_price(
                card_text
            )

            processed[
                "application_deadline"
            ] = extract_application_deadline(
                card_text
            )

            processed[
                "okpd2_codes"
            ] = extract_okpd2_codes(
                card_text
            )

            # -----------------------------------------------------
            # ОТЛАДКА ДЛЯ 223-ФЗ, ЕСЛИ СРОК НЕ НАЙДЕН
            # -----------------------------------------------------

            if (
                tender["law"] == "223-ФЗ"
                and
                processed["application_deadline"] is None
            ):

                print()
                print(
                    "========== ОТЛАДКА СРОКА 223-ФЗ =========="
                )

                # ---------------------------------------------
                # Ищем строки обычного текста карточки,
                # связанные со сроком / подачей заявок
                # ---------------------------------------------

                relevant_lines = []

                for line in card_text.splitlines():

                    line_clean = line.strip()

                    if not line_clean:
                        continue

                    line_lower = line_clean.lower()

                    if any(
                        keyword in line_lower
                        for keyword in [
                            "подач",
                            "прием",
                            "заяв",
                            "срок",
                            "оконч",
                        ]
                    ):

                        relevant_lines.append(
                            line_clean
                        )


                if relevant_lines:

                    print(
                        "\nСтроки карточки со словами "
                        "«срок / подача / заявки»:"
                    )

                    for line in relevant_lines[:30]:

                        print(
                            "  ",
                            line
                        )

                else:

                    print(
                        "\nВ обычном тексте карточки "
                        "подходящих строк не найдено."
                    )


                # ---------------------------------------------
                # Теперь смотрим исходный HTML
                # ---------------------------------------------

                try:

                    html = tender_page.content()

                    print(
                        "\nРазмер HTML карточки:",
                        len(html)
                    )


                    # Ищем участки HTML, где встречаются
                    # слова о подаче заявок.
                    #
                    # Выводим небольшой фрагмент вокруг
                    # каждого совпадения.

                    html_lower = html.lower()

                    keywords = [
                        "подач",
                        "прием заяв",
                        "окончание",
                        "срок подачи",
                    ]

                    printed_positions = set()


                    for keyword in keywords:

                        start_position = 0


                        while True:

                            position = html_lower.find(
                                keyword,
                                start_position
                            )

                            if position == -1:
                                break


                            # Не выводим один и тот же
                            # участок HTML несколько раз.

                            block_position = (
                                position // 500
                            )


                            if block_position not in printed_positions:

                                printed_positions.add(
                                    block_position
                                )


                                start = max(
                                    0,
                                    position - 500
                                )

                                end = min(
                                    len(html),
                                    position + 1000
                                )


                                fragment = html[
                                    start:end
                                ]


                                print()
                                print(
                                    "----- ФРАГМЕНТ HTML -----"
                                )

                                print(
                                    fragment
                                )

                                print(
                                    "----- КОНЕЦ ФРАГМЕНТА -----"
                                )


                            start_position = (
                                position + len(keyword)
                            )


                            if len(
                                printed_positions
                            ) >= 10:

                                break


                        if len(
                            printed_positions
                        ) >= 10:

                            break


                except Exception as e:

                    print(
                        "❌ Ошибка чтения HTML:",
                        e
                    )


                print(
                    "========== КОНЕЦ ОТЛАДКИ =========="
                )


            print()
            print(
                "Основная информация:"
            )

            print(
                "  Название:",
                processed["title"]
            )

            print(
                "  Цена:",
                processed["price"]
            )

            print(
                "  Окончание подачи:",
                processed["application_deadline"]
            )

            print(
                "  ОКПД2:",
                processed["okpd2_codes"]
            )

            # =====================================================
            # БЫСТРЫЙ ФИЛЬТР ПО НАЗВАНИЮ
            # =====================================================

            title_screen = quick_screen_title(
                processed["title"]
            )

            print()
            print("Быстрый анализ названия:")

            if title_screen["good_matches"]:
                print(
                    "  ✅ Интересные признаки:",
                    ", ".join(title_screen["good_matches"])
                )

            if title_screen["reject_matches"]:
                print(
                    "  ❌ Красные флаги:",
                    ", ".join(title_screen["reject_matches"])
                )

            if title_screen["status"] == "quick_reject":

                print()
                print("⏭ Закупка отброшена предварительным фильтром.")
                print("Причина:", title_screen["reason"])

                processed["status"] = "quick_reject"
                processed["skip_reason"] = title_screen["reason"]
                processed["quick_reject_matches"] = (
                    title_screen["reject_matches"]
                )

                processed_tenders.append(processed)

                tender_page.close()

                continue

            print()
            print("✅ Закупка проходит предварительный фильтр по названию.")


            documents_url = find_documents_url(tender_page)


            # -----------------------------------------------------
            # 5. ПРОВЕРЯЕМ СРОК ПОДАЧИ ЗАЯВОК
            # -----------------------------------------------------

            deadline = processed["application_deadline"]

            if deadline:

                try:

                    deadline_date = datetime.strptime(
                        deadline,
                        "%d.%m.%Y"
                    ).date()

                    today = datetime.now().date()

                    if deadline_date < today:

                        print()
                        print(
                            "⏭ Тендер пропущен."
                        )

                        print(
                            f"   Срок подачи заявок истёк: "
                            f"{deadline}"
                        )

                        print(
                            f"   Сегодня: "
                            f"{today.strftime('%d.%m.%Y')}"
                        )

                        processed["status"] = "expired"

                        processed["skip_reason"] = (
                            f"Срок подачи заявок истёк: {deadline}"
                        )

                        processed_tenders.append(
                            processed
                        )

                        tender_page.close()

                        continue

                except ValueError:

                    print(
                        "⚠ Не удалось разобрать дату:",
                        deadline
                    )

            # -----------------------------------------------------
            # Если срок подачи определить не удалось
            # -----------------------------------------------------

            # -----------------------------------------------------
            # ЕСЛИ СРОК ПОДАЧИ НЕ ОПРЕДЕЛИЛСЯ
            # -----------------------------------------------------

            if not deadline:

                print()
                print(
                    "⚠ Срок подачи заявок определить не удалось."
                )


                # =================================================
                # Для 223-ФЗ дополнительно проверяем год закупки
                #
                # Формат номера:
                # 326... → 2026
                # 325... → 2025
                # 319... → 2019
                # 318... → 2018
                # =================================================

                if tender["law"] == "223-ФЗ":

                    number = tender["number"]

                    try:

                        prefix = int(
                            number[:3]
                        )


                        # -----------------------------------------
                        # 300 + XX = год
                        #
                        # Например:
                        # 326 → 2026
                        # 319 → 2019
                        # 318 → 2018
                        # -----------------------------------------

                        if 300 <= prefix <= 399:

                            tender_year = (
                                2000 + prefix - 300
                            )

                            print(
                                "Год закупки по номеру:",
                                tender_year
                            )


                            current_year = datetime.now().year


                            # -----------------------------------------
                            # Если закупка старше 2024 года,
                            # при отсутствии даты считаем её
                            # архивной и пропускаем.
                            #
                            # Не трогаем 2025/2026,
                            # потому что среди них потенциально
                            # могут быть актуальные записи.
                            # -----------------------------------------

                            if tender_year <= current_year - 2:

                                print()
                                print(
                                    "⏭ Тендер пропущен."
                                )

                                print(
                                    "   Причина: старый 223-ФЗ "
                                    "без определённого срока подачи."
                                )

                                print(
                                    "   Год закупки:",
                                    tender_year
                                )

                                processed[
                                    "status"
                                ] = "old_tender"

                                processed[
                                    "skip_reason"
                                ] = (
                                    "Старая закупка 223-ФЗ, "
                                    "срок подачи не определён"
                                )

                                processed_tenders.append(
                                    processed
                                )

                                tender_page.close()

                                continue


                    except Exception as e:

                        print(
                            "⚠ Не удалось определить "
                            f"год 223-ФЗ: {e}"
                        )


                # ---------------------------------------------
                # Если закупка новая или год определить
                # не удалось — НЕ пропускаем её.
                #
                # Оставляем её для дальнейшего анализа.
                # ---------------------------------------------

                processed[
                    "status"
                ] = "deadline_unknown"

                processed[
                    "skip_reason"
                ] = (
                    "Не удалось определить срок подачи заявок"
                )

            # -----------------------------------------------------
            # 6. ИЩЕМ ВКЛАДКУ ДОКУМЕНТЫ
            # -----------------------------------------------------

            documents_url = find_documents_url(
                tender_page
            )

            if not documents_url:

                print()
                print(
                    "⚠ Ссылка на «Документы» не найдена."
                )

                processed[
                    "status"
                ] = "documents_url_not_found"

                processed_tenders.append(
                    processed
                )

                tender_page.close()

                continue


            processed[
                "documents_url"
            ] = documents_url

            print()
            print(
                "Страница документов:",
                documents_url
            )


            # -----------------------------------------------------
            # 6. ОТКРЫВАЕМ ДОКУМЕНТЫ
            # -----------------------------------------------------

            tender_page.goto(
                documents_url,
                wait_until="domcontentloaded",
                timeout=120000
            )

            tender_page.wait_for_timeout(
                4000
            )

            print(
                "✓ Страница документов открыта."
            )


            # -----------------------------------------------------
            # 7. ПОЛУЧАЕМ СПИСОК ФАЙЛОВ
            # -----------------------------------------------------

            file_list = get_document_files(
                tender_page
            )

            processed[
                "total_files"
            ] = len(file_list)

            print()
            print(
                "Найдено файлов:",
                len(file_list)
            )

            # =====================================================
            # БЫСТРЫЙ ФИЛЬТР ПО НАЗВАНИЯМ ФАЙЛОВ
            # =====================================================

            file_screen = quick_screen_files(file_list)

            print()
            print("Быстрый анализ названий файлов:")

            if file_screen["reject_matches"]:

                for match in file_screen["reject_matches"]:

                    print(
                        "  ❌",
                        match["file"],
                        "→",
                        match["keyword"]
                    )

            else:

                print(
                    "  ✅ Явных красных флагов "
                    "в названиях файлов не найдено."
                )


            if file_screen["status"] == "quick_reject":

                print()
                print(
                    "⏭ Закупка отброшена "
                    "по названиям документов."
                )

                processed["status"] = "quick_reject"
                processed["skip_reason"] = (
                    "В названиях прикрепленных документов "
                    "обнаружены признаки сложной закупки"
                )
                processed["quick_reject_matches"] = (
                    file_screen["reject_matches"]
                )

                processed_tenders.append(processed)

                tender_page.close()

                continue


            # -----------------------------------------------------
            # 8. ПОЛУЧАЕМ И ИЗВЛЕКАЕМ ТЕКСТ
            # -----------------------------------------------------

            loaded_files = []
            seen_file_hashes = set()


            for file_index, file_info in enumerate(
                file_list,
                start=1
            ):

                print()
                print(
                    "-" * 65
                )

                print(
                    f"ФАЙЛ {file_index}/{len(file_list)}"
                )

                print(
                    "Название:",
                    file_info["name"]
                )

                print(
                    "URL:",
                    file_info["url"]
                )


                try:

                    response_data = get_file_response(
                        tender_page,
                        file_info["url"]
                    )

                    if response_data is None:

                        print(
                            "❌ Не удалось получить файл."
                        )

                        continue


                    body = response_data["body"]


                    print(
                        "✓ Получено байт:",
                        len(body)
                    )


                    # ---------------------------------------------------------
                    # Проверяем, не является ли файл дубликатом
                    #
                    # Иногда ЕИС отдаёт один и тот же документ несколько раз
                    # с разными UID/ссылками.
                    # ---------------------------------------------------------

                    file_hash = hashlib.sha256(
                        body
                    ).hexdigest()


                    if file_hash in seen_file_hashes:

                        print(
                            "⏭ Дубликат файла. "
                            "Пропускаем повторную обработку."
                        )

                        continue


                    seen_file_hashes.add(
                        file_hash
                    )


                    print(
                        "✓ Файл уникальный."
                    )


                    detected_format = detect_file_format(
                        body,
                        file_info["name"]
                    )


                    print(
                        "Формат:",
                        detected_format
                    )


                    extracted_result = extract_text_from_file(
                        body,
                        file_info["name"]
                    )


                    extracted_text = extracted_result["text"]


                    print(
                        "Размер текста:",
                        len(extracted_text)
                    )


                    loaded_files.append(
                        {
                            "name": file_info["name"],
                            "url": file_info["url"],
                            "format": detected_format,
                            "size": len(body),
                            "text": extracted_text,
                            "body": body,
                        }
                    )


                    print(
                        "✓ Файл обработан."
                    )


                except Exception as e:

                    print(
                        "❌ Ошибка обработки файла:",
                        e
                    )


            processed[
                "files"
            ] = loaded_files

            processed[
                "loaded_files"
            ] = len(loaded_files)


            # -----------------------------------------------------
            # 9. СОБИРАЕМ ВЕСЬ ТЕКСТ ДОКУМЕНТОВ
            # -----------------------------------------------------

            all_document_parts = []


            for file_data in loaded_files:

                all_document_parts.append(
                    "\n".join(
                        [
                            "==================================================",
                            f"ДОКУМЕНТ: {file_data['name']}",
                            f"ФОРМАТ: {file_data['format']}",
                            "==================================================",
                            file_data["text"]
                        ]
                    )
                )


            processed[
                "documents_text"
            ] = "\n\n".join(
                all_document_parts
            )


            print()
            print(
                "Всего извлечено символов:",
                len(
                    processed["documents_text"]
                )
            )

            # ============================================================
            # ГЛУБОКИЙ ПРЕДВАРИТЕЛЬНЫЙ АНАЛИЗ ДОКУМЕНТОВ
            # ============================================================

            deep_analysis = deep_screen_documents(
                processed["documents_text"],
                loaded_files
            )

            processed["deep_analysis"] = deep_analysis

            work_profile = build_work_profile(
                processed["documents_text"],
                deep_analysis,
                loaded_files,
                processed["title"]
            )

            processed["work_profile"] = work_profile

            print()
            print("=" * 60)
            print("СТРУКТУРНЫЙ АНАЛИЗ ЗАКУПКИ")
            print("=" * 60)

            print("Типы работ:")
            if work_profile["work_types"]:
                for item in work_profile["work_types"]:
                    print(f"  • {item}")
            else:
                print("  — не определено")

            print()
            print("Результаты / материалы:")
            if work_profile["deliverables"]:
                for item in work_profile["deliverables"]:
                    print(f"  • {item}")
            else:
                print("  — не определено")

            print()
            print("Риски:")

            for name, value in work_profile["risks"].items():
                status = "ДА" if value else "НЕТ"
                print(f"  {name}: {status}")

            print()
            print("Положительные признаки:")

            for name, value in work_profile["positive_signals"].items():
                status = "ДА" if value else "НЕТ"
                print(f"  {name}: {status}")

            print()
            print("Условия выполнения:")

            if work_profile["completion_terms"]:
                for item in work_profile["completion_terms"]:
                    print(f"  → {item}")
            else:
                print("  — срок выполнения не найден")

            print()
            print("Доработки / замечания:")

            if work_profile["revision_terms"]:
                for item in work_profile["revision_terms"]:
                    print(f"  ⚠ {item}")
            else:
                print("  ✅ Явных условий об ограничении/неограниченности доработок не найдено")

            print()
            print("Согласования:")

            if work_profile["approval_terms"]:
                for item in work_profile["approval_terms"]:
                    print(f"  ⚠ {item}")
            else:
                print("  ✅ Явных требований к согласованиям не найдено")


            print()
            print("=" * 60)
            print("ГЛУБОКИЙ АНАЛИЗ ДОКУМЕНТОВ")
            print("=" * 60)


            # ------------------------------------------------------------
            # КРАСНЫЕ ФЛАГИ
            # ------------------------------------------------------------

            if deep_analysis["red_flags"]:

                print()
                print("❌ НАЙДЕНЫ ПОТЕНЦИАЛЬНЫЕ РИСКИ:")

                for category, data in (
                    deep_analysis["red_flags"].items()
                ):

                    print()
                    print(
                        "  Категория:",
                        category
                    )

                    print(
                        "  Совпадений:",
                        data["count"]
                    )

                    for snippet in data["snippets"][:2]:

                        print(
                            "  →",
                            snippet[:500]
                        )

            else:

                print()
                print(
                    "✅ Явных красных признаков "
                    "в тексте не найдено."
                )


            # ------------------------------------------------------------
            # ПОЛОЖИТЕЛЬНЫЕ ПРИЗНАКИ
            # ------------------------------------------------------------

            if deep_analysis["good_flags"]:

                print()
                print("✅ ПОЛОЖИТЕЛЬНЫЕ ПРИЗНАКИ:")

                for category, data in (
                    deep_analysis["good_flags"].items()
                ):

                    print(
                        "  ",
                        category,
                        "—",
                        data["count"],
                        "совпадений"
                    )

            else:

                print()
                print(
                    "⚠ Положительных архитектурных "
                    "признаков немного или не найдено."
                )


            # ------------------------------------------------------------
            # В КАКИХ ФАЙЛАХ ОБНАРУЖЕНЫ ПРИЗНАКИ
            # ------------------------------------------------------------

            if deep_analysis["file_findings"]:

                print()
                print("ФАЙЛЫ С НАЙДЕННЫМИ ПРИЗНАКАМИ:")

                for finding in (
                    deep_analysis["file_findings"]
                ):

                    print()
                    print(
                        "  📄",
                        finding["file"]
                    )

                    if finding["red_flags"]:

                        print(
                            "     ❌ Риски:",
                            ", ".join(
                                finding["red_flags"]
                            )
                        )

                    if finding["good_flags"]:

                        print(
                            "     ✅ Плюсы:",
                            ", ".join(
                                finding["good_flags"]
                            )
                        )



        except Exception as e:

            processed[
                "status"
            ] = "error"

            print()
            print(
                "❌ Ошибка обработки закупки:",
                e
            )

        finally:

            try:
                tender_page.close()
            except Exception:
                pass


        # ---------------------------------------------------------
        # СОХРАНЯЕМ РЕЗУЛЬТАТ В ПАМЯТИ
        # ---------------------------------------------------------

        processed_tenders.append(
            processed
        )


# ============================================================
# ФИНАЛЬНАЯ СВОДКА ТЕКУЩЕГО ЗАПУСКА
# ============================================================

print()
print("=" * 70)
print("ОБРАБОТКА ВСЕХ ЗАКУПОК ЗАВЕРШЕНА")
print("=" * 70)


# ------------------------------------------------------------
# Формируем результаты только для тендеров,
# найденных в ТЕКУЩЕМ запуске.
# ------------------------------------------------------------

processed_by_number = {}

for item in processed_tenders:

    number = item.get("number")

    if number:
        processed_by_number[number] = item


current_results = []

for tender in tenders:

    number = tender.get("number")

    if number in processed_by_number:

        current_results.append(
            processed_by_number[number]
        )


# ------------------------------------------------------------
# Статистика
# ------------------------------------------------------------

successful_count = sum(
    1
    for item in current_results
    if item.get("status") == "ok"
)


with_documents_count = sum(
    1
    for item in current_results
    if (item.get("loaded_files") or 0) > 0
)


skipped_count = sum(
    1
    for item in current_results
    if item.get("status") != "ok"
)


print()
print(
    "Всего найдено закупок:",
    len(tenders)
)

print(
    "Всего обработано:",
    len(current_results)
)

print(
    "Успешно обработано:",
    successful_count
)

print(
    "С документами:",
    with_documents_count
)

print(
    "Пропущено / отфильтровано:",
    skipped_count
)


# ------------------------------------------------------------
# Проверка: все ли найденные закупки попали в результаты
# ------------------------------------------------------------

if len(current_results) != len(tenders):

    print()
    print("⚠ ВНИМАНИЕ")

    print(
        "Найдено закупок:",
        len(tenders)
    )

    print(
        "Есть результатов обработки:",
        len(current_results)
    )

    print(
        "Разница:",
        len(tenders) - len(current_results)
    )


# ============================================================
# КРАТКАЯ СВОДКА
# ============================================================

print()
print("=" * 70)
print("КРАТКАЯ СВОДКА")
print("=" * 70)


for index, item in enumerate(
    current_results,
    start=1
):

    number = item.get(
        "number",
        "неизвестно"
    )

    law = item.get(
        "law",
        "неизвестно"
    )

    title = item.get(
        "title",
        ""
    )

    price = item.get(
        "price",
        ""
    )

    deadline = item.get(
        "application_deadline"
    )

    okpd2 = item.get(
        "okpd2_codes",
        []
    )

    total_files = (
        item.get("total_files")
        or 0
    )

    loaded_files = (
        item.get("loaded_files")
        or 0
    )

    documents_text = (
        item.get("documents_text")
        or ""
    )

    status = item.get(
        "status",
        "unknown"
    )

    skip_reason = item.get(
        "skip_reason"
    )


    print()
    print(
        f"{index}. {number} ({law})"
    )

    print(
        "   Название:",
        title
    )

    print(
        "   Цена:",
        price
    )

    print(
        "   Срок:",
        deadline
    )

    print(
        "   ОКПД2:",
        ", ".join(okpd2)
        if okpd2
        else "не найден"
    )

    print(
        "   Статус:",
        status
    )

    print(
        "   Документов:",
        loaded_files,
        "/",
        total_files
    )

    print(
        "   Текст документов:",
        len(documents_text),
        "симв."
    )

    if skip_reason:

        print(
            "   Причина:",
            skip_reason
        )


# ============================================================
# СТРУКТУРА ПЕРВОГО ТЕНДЕРА
# ============================================================

print()
print("=" * 70)
print("СТРУКТУРА ПЕРВОГО ТЕНДЕРА")
print("=" * 70)


if current_results:

    first_tender = next(
    (
        item
        for item in current_results
        if item.get("deep_analysis")
    ),
    None
    )

    print(
        json.dumps(
            {
                "number": first_tender.get(
                    "number"
                ),

                "law": first_tender.get(
                    "law"
                ),

                "title": first_tender.get(
                    "title"
                ),

                "price": first_tender.get(
                    "price"
                ),

                "application_deadline": first_tender.get(
                    "application_deadline"
                ),

                "okpd2_codes": first_tender.get(
                    "okpd2_codes"
                ),

                "tender_url": first_tender.get(
                    "tender_url"
                ),

                "documents_url": first_tender.get(
                    "documents_url"
                ),

                "total_files": first_tender.get(
                    "total_files",
                    0
                ),

                "loaded_files": first_tender.get(
                    "loaded_files",
                    0
                ),

                "documents_text_length": len(
                    first_tender.get(
                        "documents_text",
                        ""
                    )
                    or ""
                ),

                "deep_analysis": first_tender.get(
                    "deep_analysis",
                    {}
                ),
            },
            ensure_ascii=False,
            indent=2
        )
    )

else:

    print(
        "Нет результатов для отображения."
    )


    # =========================================================
    # ЗАВЕРШЕНИЕ
    # =========================================================

    print()
    print("=" * 70)
    print("РАБОТА ЗАКОНЧЕНА")
    print("=" * 70)

    print()
    print(
        "Браузер оставлен открытым."
    )

    input(
        "Нажми Enter в PowerShell для закрытия..."
    )

    browser.close()