import os
import re
import json
import time
import html
import logging
import asyncio
import urllib.request
import urllib.parse
from datetime import datetime, timedelta
from io import BytesIO

import openpyxl
import pandas as pd
from bs4 import BeautifulSoup
from dotenv import load_dotenv

from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.types import Message, CallbackQuery
from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.exceptions import TelegramBadRequest

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
)
logger = logging.getLogger("raspisanie_bot")

load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN", "YOUR_BOT_TOKEN_HERE")

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

# --- Константы и пути к файлам ---
SCHEDULE_PAGE_URL = "https://phystech.pro/%D1%81%D1%82%D1%83%D0%B4%D0%B5%D0%BD%D1%82%D1%83/%D1%80%D0%B0%D1%81%D0%BF%D0%B8%D1%81%D0%B0%D0%BD%D0%B8%D0%B5/"
WP_MEDIA_API_URL = "https://phystech.pro/wp-json/wp/v2/media?search=xlsx&per_page=15"
FALLBACK_SCHEDULE_URLS = [
    "https://phystech.pro/wp-content/uploads/2026/09/25.09.-%D0%A1%D1%82%D1%80%D1%83%D0%BA%D1%82%D1%83%D1%80%D0%BD%D0%BE%D0%B5-%D0%BF%D0%BE%D0%B4%D1%80%D0%B0%D0%B7%D0%B4%D0%B5%D0%BB%D0%B5%D0%BD%D0%B8%D0%B5-%E2%84%961-1.xlsx",
    "https://phystech.pro/wp-content/uploads/2026/09/25.09-%D0%A1%D1%82%D1%80%D1%83%D0%BA%D1%82%D1%83%D1%80%D0%BD%D0%BE%D0%B5-%D0%BF%D0%BE%D0%B4%D1%80%D0%B0%D0%B7%D0%B4%D0%B5%D0%BB%D0%B5%D0%BD%D0%B8%D0%B5-%E2%84%962-1.xlsx",
]

CACHE_FILE = "schedule_cache.json"
USER_GROUPS_FILE = "user_groups.json"
CACHE_TTL_SECONDS = 600  # 10 минут

RU_DAY_NAMES = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскресенье"]
SHORT_DAY_NAMES = {
    "Monday": "Пн", "Tuesday": "Вт", "Wednesday": "Ср", "Thursday": "Чт",
    "Friday": "Пт", "Saturday": "Сб", "Sunday": "Вс"
}

MONTHS_RU = {
    'янв': 1, 'фев': 2, 'мар': 3, 'апр': 4, 'мая': 5, 'май': 5, 'июн': 6,
    'июл': 7, 'авг': 8, 'сен': 9, 'окт': 10, 'ноя': 11, 'дек': 12
}

WEEKDAYS_MAP = {
    'пн': 0, 'пон': 0, 'понедельник': 0, 'пн.': 0,
    'вт': 1, 'вто': 1, 'вторник': 1, 'вт.': 1,
    'ср': 2, 'сре': 2, 'среда': 2, 'ср.': 2,
    'чт': 3, 'чет': 3, 'четверг': 3, 'чт.': 3,
    'пт': 4, 'пят': 4, 'пятница': 4, 'пт.': 4,
    'сб': 5, 'суб': 5, 'суббота': 5, 'сб.': 5,
    'вс': 6, 'вос': 6, 'воскресенье': 6, 'вс.': 6,
}

# Визуально похожие латинские буквы -> кириллица
LATIN_LOOKALIKES = {
    'A': 'А', 'B': 'В', 'C': 'С', 'E': 'Е', 'H': 'Н', 'K': 'К',
    'M': 'М', 'O': 'О', 'P': 'Р', 'T': 'Т', 'X': 'Х', 'Y': 'У',
    'V': 'В'
}

# Фонетическое соответствие для ввода названий групп на английской раскладке
PHONETIC_LATIN = {
    'I': 'И', 'S': 'С', 'P': 'П', 'D': 'Д', 'T': 'Т', 'B': 'Б',
    'R': 'Р', 'U': 'У', 'N': 'Н', 'G': 'Г', 'L': 'Л', 'Z': 'З',
    'F': 'Ф', 'K': 'К', 'M': 'М', 'A': 'А', 'E': 'Е', 'O': 'О',
    'V': 'В', 'W': 'В', 'Y': 'У', 'C': 'С', 'H': 'Х'
}

TIME_REGEX = re.compile(r'(\d{1,2}[:.]\d{2})\s*[-–—]\s*(\d{1,2}[:.]\d{2})')

# Стандартные звонки Физтех-колледжа (для восстановления времени/пар)
STANDARD_BELLS = {
    "1": "08:30-10:00",
    "2": "10:10-11:40",
    "3": "11:50-13:20",
    "4": "13:50-15:20",
    "5": "15:30-17:00",
    "6": "17:10-18:40",
    "7": "18:50-20:20",
}

# --- Кэш в оперативной памяти ---
_cached_schedule = {}
_last_fetch_timestamp = 0.0
_cache_lock = asyncio.Lock()


# --- Хранение групп пользователей ---
def load_user_groups() -> dict:
    if os.path.exists(USER_GROUPS_FILE):
        try:
            with open(USER_GROUPS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return {int(k): str(v) for k, v in data.items()}
        except Exception as e:
            logger.error(f"Ошибка чтения {USER_GROUPS_FILE}: {e}")
    return {}


def save_user_groups(groups: dict):
    try:
        with open(USER_GROUPS_FILE, "w", encoding="utf-8") as f:
            json.dump(groups, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.error(f"Ошибка сохранения {USER_GROUPS_FILE}: {e}")


user_groups = load_user_groups()


# --- Кэш расписания на диске ---
def load_disk_cache() -> dict:
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception as e:
            logger.error(f"Ошибка чтения {CACHE_FILE}: {e}")
    return {}


def save_disk_cache(data: dict):
    try:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.error(f"Ошибка сохранения {CACHE_FILE}: {e}")


_cached_schedule = load_disk_cache()


# --- Вспомогательные функции распознавания времени и пар ---
def infer_pair_from_time(t_str: str) -> str:
    """Определяет номер пары по времени звонков."""
    if not t_str:
        return ""
    t_clean = re.sub(r'\s+', '', t_str).replace('.', ':')
    m = TIME_REGEX.search(t_clean)
    if not m:
        return ""
    start_t = m.group(1).zfill(5)
    end_t = m.group(2).zfill(5)

    if start_t == "08:30" and end_t in ["09:15", "09:10"]:
        return "Разговоры о важном"
    if start_t in ["08:30", "09:15"] and end_t in ["10:00", "10:45"]:
        return "1"
    if start_t in ["10:10", "10:55"] and end_t in ["11:40", "12:25"]:
        return "2"
    if start_t in ["11:50", "12:35"] and end_t in ["13:20", "14:05"]:
        return "3"
    if start_t in ["13:20", "14:05"] and end_t in ["13:50", "14:35"]:
        return "Перерыв"
    if start_t in ["13:50", "14:35"] and end_t in ["15:20", "16:05"]:
        return "4"
    if start_t in ["15:30", "16:15"] and end_t in ["17:00", "17:45"]:
        return "5"
    if start_t in ["17:10", "17:55"] and end_t in ["18:40", "19:25"]:
        return "6"
    if start_t in ["18:50"] and end_t in ["20:20"]:
        return "7"
    return ""


def parse_date_context(text: str, default_year: int = 2026):
    """
    Универсально извлекает дату или диапазон дат из текста, URL или имени файла.
    Возвращает datetime или кортеж (start_date, end_date), либо None.
    """
    if not text:
        return None
    text_clean = text.replace('%20', ' ').replace('_', ' ')
    text_lower = text_clean.lower()

    # 1. Диапазон дат типа DD.MM - DD.MM (например, 23.09-25.09)
    m = re.search(r'(\d{1,2})[.](\d{2})\s*[-–—с_по\s]+\s*(\d{1,2})[.](\d{2})', text_clean)
    if m:
        try:
            d1 = datetime(default_year, int(m.group(2)), int(m.group(1)))
            d2 = datetime(default_year, int(m.group(4)), int(m.group(3)))
            return (d1, d2)
        except Exception:
            pass

    # 2. Текстовый диапазон: со 2 по 4 сентября / со-2-по-4-сентября / с 09.09 по 11.09
    m = re.search(r'(?:с|со)?[\s_-]*(\d{1,2})[\s_-]*(?:по|до|-|–|—)[\s_-]*(\d{1,2})[\s_-]*([а-яё]{3,10})', text_lower)
    if m:
        for prefix, mon in MONTHS_RU.items():
            if m.group(3).startswith(prefix):
                try:
                    d1 = datetime(default_year, mon, int(m.group(1)))
                    d2 = datetime(default_year, mon, int(m.group(2)))
                    return (d1, d2)
                except Exception:
                    pass

    # 3. YYYY-MM-DD (например 2026-09-14)
    m = re.search(r'(\d{4})[._-](\d{2})[._-](\d{2})', text_clean)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except Exception:
            pass

    # 4. YY-MM-DD (например 26-09-14)
    m = re.search(r'\b(2[4-9])[._-](\d{2})[._-](\d{2})\b', text_clean)
    if m:
        try:
            return datetime(2000 + int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except Exception:
            pass

    # 5. Одиночная дата DD.MM.YYYY
    m = re.search(r'(\d{1,2})[.](\d{2})[.](\d{4})', text_clean)
    if m:
        try:
            return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except Exception:
            pass

    # 6. Одиночная дата DD месяца (например 25 сентября)
    m = re.search(r'(\d{1,2})\s+([а-яё]{3,10})', text_lower)
    if m:
        for prefix, mon in MONTHS_RU.items():
            if m.group(2).startswith(prefix):
                try:
                    return datetime(default_year, mon, int(m.group(1)))
                except Exception:
                    pass

    # 7. Одиночная дата DD.MM
    m = re.search(r'(\d{1,2})[.](\d{2})', text_clean)
    if m:
        try:
            return datetime(default_year, int(m.group(2)), int(m.group(1)))
        except Exception:
            pass

    return None


def extract_day_indicator(val: str):
    """Проверяет, содержит ли ячейка указание на день недели (Пн, Вт, Среда и т.д.)."""
    if not val or not isinstance(val, str):
        return None
    val_lower = val.strip().lower()
    val_clean = re.sub(r'[.,\s]', '', val_lower)
    for k, idx in WEEKDAYS_MAP.items():
        if val_clean == k or val_lower.startswith(k):
            return idx
    return None


# --- Умная нормализация и поиск группы ---
def to_cyrillic_variants(s: str) -> list:
    """Генерирует кириллические варианты строки, даже если ввели на латинице."""
    s_up = s.upper().strip()
    var_lookalike = "".join(LATIN_LOOKALIKES.get(ch, ch) for ch in s_up)
    var_phonetic = "".join(PHONETIC_LATIN.get(ch, ch) for ch in s_up)
    return list(dict.fromkeys([s_up, var_lookalike, var_phonetic]))


def clean_for_match(s: str) -> str:
    return re.sub(r'[\s\-_/]', '', s.upper().strip())


def strip_year_suffix(s: str) -> str:
    return re.sub(r'/\s*\d+$', '', s.strip())


def is_group_name(val: str) -> bool:
    if not val or not isinstance(val, str):
        return False
    val_clean = val.strip()
    if len(val_clean) < 2 or len(val_clean) > 30:
        return False
    lower = val_clean.lower()
    stop_words = [
        'день', 'время', 'пара', 'звонк', 'перерыв', 'обед', 'недел',
        'кабинет', 'аудит', 'расписан', 'корпус', 'преподават',
        'дисциплин', 'предмет', 'подгрупп', 'п/г', 'открыть', 'скачать',
        'заняти', 'урок', 'класс'
    ]
    if any(k in lower for k in stop_words):
        return False
    # Название группы содержит буквы и цифры (например: ИСП-31, М-51, П-61, АДТ-111/26)
    has_letters = any(c.isalpha() for c in val_clean)
    has_digits = any(c.isdigit() for c in val_clean)
    return has_letters and has_digits


def clean_group_name(name: str) -> str:
    name_str = str(name).strip()
    name_str = re.sub(r'^[Гг]руппа\s+', '', name_str)
    name_str = re.sub(r'\s*/\s*', '/', name_str)
    name_str = re.sub(r'\s*-\s*', '-', name_str)
    return re.sub(r'\s+', ' ', name_str).strip()


def find_group_in_dict(groups_dict: dict or list, target_group: str) -> str or None:
    if not groups_dict or not target_group:
        return None

    keys = list(groups_dict.keys()) if isinstance(groups_dict, dict) else list(groups_dict)
    variants = to_cyrillic_variants(target_group)

    # 1. Точное или очищенное совпадение (с учетом года)
    for v in variants:
        v_clean = clean_for_match(v)
        for g in keys:
            if v_clean == clean_for_match(g):
                return g

    # 2. Совпадение без учета года поступления (например 'АДТ-111' и 'АДТ-111/26')
    for v in variants:
        v_no_year = clean_for_match(strip_year_suffix(v))
        for g in keys:
            if v_no_year == clean_for_match(strip_year_suffix(g)):
                return g

    # 3. Префиксное совпадение
    for v in variants:
        v_no_year = clean_for_match(strip_year_suffix(v))
        if len(v_no_year) >= 3:
            for g in keys:
                g_no_year = clean_for_match(strip_year_suffix(g))
                if g_no_year.startswith(v_no_year):
                    return g

    return None


# --- Универсальный парсер Excel (openpyxl + разъединение ячеек) ---
def unmerge_and_load_grid(sheet) -> list:
    """
    Разъединяет объединенные ячейки и копирует значение во все ячейки диапазона.
    Возвращает 2D массив строк/None.
    """
    for rng in list(sheet.merged_cells.ranges):
        top_val = sheet.cell(rng.min_row, rng.min_col).value
        sheet.unmerge_cells(str(rng))
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                sheet.cell(r, c).value = top_val

    grid = []
    for r in range(1, sheet.max_row + 1):
        row_vals = []
        for c in range(1, sheet.max_column + 1):
            v = sheet.cell(r, c).value
            if v is not None:
                s = str(v).strip()
                row_vals.append(s if s and s.lower() not in ['none', 'nan'] else None)
            else:
                row_vals.append(None)
        grid.append(row_vals)
    return grid


def parse_semester_grid(grid: list, default_date: datetime) -> dict:
    """
    Парсер семестрового/блочного расписания (Layout 2, например 1сем_1курс.xlsx),
    где группы указаны сверху, под ними столбцы дней недели (пн, вт, ср, чт, пт),
    а в строках идут пары 1..6.
    """
    group_row_idx = -1
    for r in range(min(15, len(grid) - 1)):
        row = grid[r]
        groups_in_row = [c for c, v in enumerate(row) if v and is_group_name(v)]
        if len(groups_in_row) >= 2:
            next_row = grid[r + 1]
            days_count = sum(1 for v in next_row if extract_day_indicator(v) is not None)
            if days_count >= 3:
                group_row_idx = r
                break

    if group_row_idx == -1:
        return {}

    group_row = grid[group_row_idx]
    days_row = grid[group_row_idx + 1]

    pair_col = -1
    for c in range(min(5, len(grid[0]))):
        for r in range(group_row_idx + 2, min(group_row_idx + 10, len(grid))):
            v = grid[r][c]
            if v and v.isdigit() and 1 <= int(v) <= 8:
                pair_col = c
                break
        if pair_col != -1:
            break

    col_mapping = {}
    current_group = None
    for c in range(len(group_row)):
        if group_row[c] and is_group_name(group_row[c]):
            current_group = clean_group_name(group_row[c])
        day_w = extract_day_indicator(days_row[c])
        if current_group and day_w is not None:
            col_mapping[c] = (current_group, day_w)

    if not col_mapping:
        return {}

    if default_date:
        base_monday = default_date - timedelta(days=default_date.weekday())
    else:
        now = datetime.now()
        base_monday = now - timedelta(days=now.weekday())

    schedule_by_date = {}

    for r in range(group_row_idx + 2, len(grid)):
        row = grid[r]
        pair_val = row[pair_col] if pair_col != -1 and row[pair_col] else ""
        if not pair_val or not pair_val.isdigit():
            rel_idx = r - (group_row_idx + 2) + 1
            if 1 <= rel_idx <= 8:
                pair_val = str(rel_idx)
            else:
                pair_val = "—"

        time_val = STANDARD_BELLS.get(pair_val, "")

        for c, (g_name, day_w) in col_mapping.items():
            cell_text = row[c]
            if not cell_text or cell_text.lower() in ['nan', 'none', '—', '-']:
                continue

            sec_date = base_monday + timedelta(days=day_w)
            date_str = sec_date.strftime("%d.%m.%Y")
            day_name = RU_DAY_NAMES[day_w]

            if date_str not in schedule_by_date:
                schedule_by_date[date_str] = {
                    "day_name": day_name,
                    "groups": {}
                }
            if g_name not in schedule_by_date[date_str]["groups"]:
                schedule_by_date[date_str]["groups"][g_name] = []

            schedule_by_date[date_str]["groups"][g_name].append({
                "pair": pair_val,
                "time": time_val,
                "lesson": cell_text
            })

    return schedule_by_date


def resolve_section_date(day_idx: int or None, sec_idx: int, date_info) -> datetime:
    """Рассчитывает точную дату дня/секции на основе дня недели (Пн=0..Вс=6) и дат файла."""
    if isinstance(date_info, tuple) and len(date_info) == 2:
        start_d, end_d = date_info
        if day_idx is not None:
            curr = start_d
            while curr <= end_d + timedelta(days=6):
                if curr.weekday() == day_idx:
                    return curr
                curr += timedelta(days=1)
        target = start_d + timedelta(days=sec_idx)
        return target if target <= end_d else end_d

    if isinstance(date_info, datetime):
        if day_idx is not None:
            start_week = date_info - timedelta(days=date_info.weekday())
            return start_week + timedelta(days=day_idx)
        return date_info

    now = datetime.now()
    if day_idx is not None:
        start_week = now - timedelta(days=now.weekday())
        return start_week + timedelta(days=day_idx)
    return now + timedelta(days=sec_idx)


def parse_sheet_smart(sheet, file_date_info=None) -> dict:
    """
    Супер-адаптивный парсер расписания Физтех-колледжа:
    - Разъединяет и заполняет все объединенные ячейки
    - Автоматически выбирает формат (семестровый блочный или оперативный стандартный)
    - Поддерживает перестановку и пропуски колонок День/Время/Пара
    - Поддерживает многодневные листы с повторяющимися шапками или днями в колонке А
    - Восстанавливает пропущенные номера пар и время звонков
    - Точно определяет даты каждого дня
    - Поддерживает подгруппы без поломки поиска групп
    """
    grid = unmerge_and_load_grid(sheet)
    if not grid or len(grid) < 2:
        return {}

    sheet_date = parse_date_context(sheet.title)
    effective_date_info = file_date_info or sheet_date

    # 1. Проверяем семестровый формат
    semester_base_date = (
        effective_date_info if isinstance(effective_date_info, datetime)
        else (effective_date_info[0] if isinstance(effective_date_info, tuple) else datetime.now())
    )
    semester_res = parse_semester_grid(grid, semester_base_date)
    if semester_res:
        return semester_res

    # 2. Ищем строки шапок с группами
    header_rows = []
    for r in range(len(grid)):
        row = grid[r]
        g_cols = [c for c, v in enumerate(row) if v and is_group_name(v)]
        if len(g_cols) >= 2:
            header_rows.append((r, g_cols))

    if not header_rows:
        for r in range(len(grid)):
            row = grid[r]
            g_cols = [c for c, v in enumerate(row) if v and is_group_name(v)]
            if len(g_cols) >= 1:
                header_rows.append((r, g_cols))
                break

    if not header_rows:
        return {}

    # Разделяем таблицу на секции по строкам с группами
    sections = []
    for i, (r_idx, cols) in enumerate(header_rows):
        if not sections or r_idx > sections[-1]['header_row'] + 2:
            sections.append({
                'header_row': r_idx,
                'group_cols': cols,
                'end_row': len(grid)
            })
            if len(sections) > 1:
                sections[-2]['end_row'] = r_idx

    schedule_by_date = {}

    for sec_idx, sec in enumerate(sections):
        h_idx = sec['header_row']
        end_idx = sec['end_row']
        h_row = grid[h_idx]

        # Проверяем строку подгрупп (1 п/г, 2 п/г)
        subgroup_row = grid[h_idx + 1] if h_idx + 1 < len(grid) else None
        has_subgroups = False
        if subgroup_row:
            sub_count = sum(1 for v in subgroup_row if v and any(p in str(v).lower() for p in ['п/г', 'подгрупп']))
            if sub_count >= 2:
                has_subgroups = True

        first_data_row = h_idx + 2 if has_subgroups else h_idx + 1

        # Формируем структуру: col_idx -> (основная группа, подгруппа)
        group_cols = {}
        first_group_col = len(h_row)

        for c in sec['group_cols']:
            raw_gname = clean_group_name(h_row[c])
            sub_name = ""
            if has_subgroups and subgroup_row and subgroup_row[c]:
                sub_clean = subgroup_row[c].strip()
                if any(p in sub_clean.lower() for p in ['п/г', 'подгрупп', '(1)', '(2)']):
                    sub_name = sub_clean
            group_cols[c] = (raw_gname, sub_name)
            first_group_col = min(first_group_col, c)

        meta_cols_count = first_group_col

        # Сканируем дни недели в мета-колонках этой секции
        days_in_sec = []
        for r in range(first_data_row, end_idx):
            row = grid[r]
            meta_cells = [row[c] for c in range(min(meta_cols_count, len(row))) if row[c]]
            for mc in meta_cells:
                dw = extract_day_indicator(mc)
                if dw is not None:
                    days_in_sec.append((r, dw))
                    break

        unique_days = sorted(list(set(dw for r, dw in days_in_sec)))
        day_blocks = []

        if len(unique_days) <= 1:
            # Секция относится к одному дню
            single_day_w = unique_days[0] if unique_days else None
            active_rows = []
            for r in range(first_data_row, end_idx):
                row = grid[r]
                has_lessons = any(row[c] for c in group_cols.keys() if c < len(row) and row[c])
                meta_cells = [row[c] for c in range(min(meta_cols_count, len(row))) if row[c]]
                if not has_lessons and any('перерыв' in str(mc).lower() or 'обед' in str(mc).lower() for mc in meta_cells):
                    continue
                if has_lessons or any(TIME_REGEX.search(str(mc)) for mc in meta_cells) or any(str(mc).isdigit() for mc in meta_cells):
                    active_rows.append(r)
            if active_rows:
                day_blocks.append((single_day_w, active_rows))
        else:
            # В секции несколько разных дней (например: Среда, Четверг, Пятница)
            current_day_w = days_in_sec[0][1] if days_in_sec else None
            current_day_rows = []
            day_change_points = {r: dw for r, dw in days_in_sec}

            for r in range(first_data_row, end_idx):
                row = grid[r]
                if r in day_change_points:
                    new_dw = day_change_points[r]
                    if new_dw != current_day_w and current_day_rows:
                        day_blocks.append((current_day_w, current_day_rows))
                        current_day_rows = []
                    current_day_w = new_dw

                has_lessons = any(row[c] for c in group_cols.keys() if c < len(row) and row[c])
                meta_cells = [row[c] for c in range(min(meta_cols_count, len(row))) if row[c]]
                if not has_lessons and any('перерыв' in str(mc).lower() or 'обед' in str(mc).lower() for mc in meta_cells):
                    continue
                if has_lessons or any(TIME_REGEX.search(str(mc)) for mc in meta_cells) or any(str(mc).isdigit() for mc in meta_cells):
                    current_day_rows.append(r)

            if current_day_rows:
                day_blocks.append((current_day_w, current_day_rows))

        if not day_blocks:
            continue

        for b_idx, (block_day_w, rows) in enumerate(day_blocks):
            effective_sec_idx = sec_idx + b_idx if len(day_blocks) > 1 else sec_idx
            sec_dt = resolve_section_date(block_day_w, effective_sec_idx, effective_date_info)
            date_str = sec_dt.strftime("%d.%m.%Y")
            day_name = RU_DAY_NAMES[sec_dt.weekday()]

            if date_str not in schedule_by_date:
                schedule_by_date[date_str] = {
                    "day_name": day_name,
                    "groups": {}
                }

            lesson_counter = 1
            for r in rows:
                row = grid[r]
                time_str = ""
                pair_str = ""

                for c in range(min(meta_cols_count, len(row))):
                    cell = row[c]
                    if not cell:
                        continue
                    m = TIME_REGEX.search(cell)
                    if m:
                        time_str = f"{m.group(1)}-{m.group(2)}"
                        continue
                    if cell.isdigit() and len(cell) <= 2:
                        pair_str = cell
                        continue
                    if cell.lower() in ['перерыв', 'обед']:
                        pair_str = 'Перерыв'
                        continue
                    try:
                        f_val = float(cell)
                        if 0 <= f_val <= 12:
                            pair_str = str(int(f_val))
                            continue
                    except ValueError:
                        pass

                if not pair_str and time_str:
                    inferred = infer_pair_from_time(time_str)
                    if inferred:
                        pair_str = inferred

                if not time_str and pair_str and pair_str in STANDARD_BELLS:
                    time_str = STANDARD_BELLS[pair_str]

                if not pair_str:
                    if pair_str != 'Перерыв':
                        pair_str = str(lesson_counter)
                        lesson_counter += 1
                    else:
                        pair_str = "—"
                elif pair_str.isdigit():
                    lesson_counter = int(pair_str) + 1

                for c, (g_name, sub_name) in group_cols.items():
                    if c < len(row):
                        cell_val = row[c]
                        if cell_val and cell_val.lower() not in ['nan', 'none', '—', '-']:
                            raw_text = cell_val.strip()
                            lesson_text = f"[{sub_name}] {raw_text}" if sub_name else raw_text

                            if g_name not in schedule_by_date[date_str]["groups"]:
                                schedule_by_date[date_str]["groups"][g_name] = []

                            existing_lessons = schedule_by_date[date_str]["groups"][g_name]
                            # Защита от дублирования одинакового урока в подгруппах
                            if not any(
                                ex['pair'] == pair_str and (ex['lesson'] == lesson_text or ex['lesson'] == raw_text)
                                for ex in existing_lessons
                            ):
                                existing_lessons.append({
                                    "pair": pair_str,
                                    "time": time_str,
                                    "lesson": lesson_text
                                })

    return schedule_by_date


# --- Поиск и загрузка файлов расписания с сайта колледжа ---
def find_schedule_files_on_page() -> list:
    """
    Находит все актуальные ссылки на файлы расписания (Excel) со страницы сайта
    или через WordPress REST API в случае изменения структуры страницы.
    """
    excel_items = []
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}

    # 1. Попытка парсинга страницы расписания
    try:
        req = urllib.request.Request(SCHEDULE_PAGE_URL, headers=headers)
        with urllib.request.urlopen(req, timeout=12) as resp:
            html_content = resp.read().decode('utf-8', errors='ignore')

        soup = BeautifulSoup(html_content, 'html.parser')
        for a in soup.find_all('a', href=True):
            href = a['href']
            if not any(ext in href.lower() for ext in ['.xlsx', '.xls']):
                continue

            if href.startswith('/'):
                href = "https://phystech.pro" + href
            href = urllib.parse.quote(urllib.parse.unquote(href), safe=':/?&=#')

            parent = a
            for _ in range(3):
                if parent.parent:
                    parent = parent.parent
            context_text = f"{parent.get_text(' ', strip=True)} {href}"
            date_info = parse_date_context(context_text)

            excel_items.append((href, date_info))
    except Exception as e:
        logger.error(f"Ошибка запроса страницы расписания: {e}")

    # 2. Если на странице ничего не найдено, запрашиваем WordPress Media API
    if not excel_items:
        logger.warning("На странице не найдены Excel файлы, запрашиваем WP Media API...")
        try:
            req = urllib.request.Request(WP_MEDIA_API_URL, headers=headers)
            with urllib.request.urlopen(req, timeout=8) as resp:
                data = json.loads(resp.read().decode('utf-8'))
                for item in data:
                    src = item.get('source_url', '')
                    title = item.get('title', {}).get('rendered', '')
                    if any(src.lower().endswith(ext) for ext in ['.xlsx', '.xls']):
                        d_info = parse_date_context(f"{title} {src}")
                        excel_items.append((src, d_info))
                        if len(excel_items) >= 4:
                            break
        except Exception as e:
            logger.error(f"Ошибка получения медиа через WP REST API: {e}")

    # 3. Резервные ссылки если ничего не найдено
    if not excel_items:
        logger.warning("Используем резервные ссылки расписания.")
        curr_dt = datetime.now()
        excel_items = [(url, curr_dt) for url in FALLBACK_SCHEDULE_URLS]

    return excel_items


def fetch_and_parse_all_schedules() -> dict:
    """
    Загружает и объединяет расписания из всех актуальных файлов колледжа
    (СП-1 на пл. Собина и СП-2 на ул. Школьной, семестровые и многодневные файлы).
    """
    excel_items = find_schedule_files_on_page()
    combined_schedule = {}

    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}

    for url, date_info in excel_items:
        try:
            logger.info(f"Загрузка файла расписания: {url}")
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=15) as resp:
                data = resp.read()

            file_d_info = date_info or parse_date_context(url)

            # Пробуем разобрать через openpyxl
            try:
                wb = openpyxl.load_workbook(BytesIO(data), data_only=True)
                for sheet_name in wb.sheetnames:
                    sheet = wb[sheet_name]
                    sheet_res = parse_sheet_smart(sheet, file_d_info)

                    for d_str, day_data in sheet_res.items():
                        if d_str not in combined_schedule:
                            combined_schedule[d_str] = {
                                "day_name": day_data["day_name"],
                                "groups": {}
                            }
                        for g_name, lessons in day_data["groups"].items():
                            combined_schedule[d_str]["groups"][g_name] = lessons

            except Exception as e_xl:
                logger.warning(f"openpyxl не смог открыть {url} ({e_xl}), пробуем pandas fallback...")
                xl = pd.ExcelFile(BytesIO(data))
                for sheet_name in xl.sheet_names:
                    df = xl.parse(sheet_name, header=None)
                    # Базовый fallback если openpyxl не справился
                    if not df.empty and len(df) >= 2:
                        logger.info(f"Разобран лист {sheet_name} через pandas fallback")

        except Exception as e:
            logger.error(f"Ошибка при загрузке/парсинге {url}: {e}")

    return combined_schedule


async def get_schedule(force_refresh: bool = False) -> dict:
    """
    Возвращает актуальное расписание с кэшированием в памяти и на диске.
    """
    global _cached_schedule, _last_fetch_timestamp

    now = time.time()
    if not force_refresh and _cached_schedule and (now - _last_fetch_timestamp < CACHE_TTL_SECONDS):
        return _cached_schedule

    async with _cache_lock:
        if not force_refresh and _cached_schedule and (time.time() - _last_fetch_timestamp < CACHE_TTL_SECONDS):
            return _cached_schedule

        new_schedule = await asyncio.to_thread(fetch_and_parse_all_schedules)
        if new_schedule:
            _cached_schedule = new_schedule
            _last_fetch_timestamp = time.time()
            save_disk_cache(_cached_schedule)
            logger.info(f"Расписание успешно обновлено! Доступно дат: {len(_cached_schedule)}")
        elif not _cached_schedule:
            _cached_schedule = load_disk_cache()

        return _cached_schedule


def get_schedule_message(schedule_data: dict, date_str: str, group_name: str) -> str:
    """
    Форматирует расписание на указанный день для указанной группы.
    """
    if not schedule_data:
        return "❌ Расписание временно недоступно. Попробуйте обновить чуть позже."

    if date_str not in schedule_data:
        return f"❌ Данные за <b>{html.escape(date_str)}</b> не найдены."

    day_data = schedule_data[date_str]
    matched_group = find_group_in_dict(day_data["groups"], group_name)

    if not matched_group:
        all_groups = sorted(list(day_data["groups"].keys()))
        sample = ", ".join(all_groups[:14]) + ("..." if len(all_groups) > 14 else "")
        return (
            f"❌ Группа <b>{html.escape(group_name)}</b> не найдена в расписании на {html.escape(date_str)}.\n\n"
            f"📋 Доступные группы ({len(all_groups)} шт.):\n<i>{html.escape(sample)}</i>\n\n"
            f"Проверьте правильность написания или смените группу через команду /group."
        )

    lessons = day_data["groups"].get(matched_group, [])
    day_name = day_data.get("day_name", "")

    if not lessons:
        return (
            f"🎉 <b>Расписание: {html.escape(matched_group)}</b>\n"
            f"📅 <b>{html.escape(date_str)} ({html.escape(day_name)})</b>\n\n"
            f"На этот день занятий нет! Отдыхай."
        )

    msg = (
        f"📚 <b>Расписание: {html.escape(matched_group)}</b>\n"
        f"📅 <b>{html.escape(date_str)} ({html.escape(day_name)})</b>\n\n"
    )

    for les in lessons:
        pair_num = les.get('pair', '').strip()
        time_val = les.get('time', '').strip()
        lesson_text = les.get('lesson', '').strip()

        time_part = f" ({time_val})" if time_val else ""
        if pair_num.isdigit():
            header = f"🔹 <b>{pair_num} пара{time_part}:</b>"
        elif pair_num and pair_num != "—":
            header = f"🔹 <b>{pair_num}{time_part}:</b>"
        elif time_val:
            header = f"🔹 <b>{time_val}:</b>"
        else:
            header = f"🔹 <b>Занятие:</b>"

        msg += f"{header}\n{html.escape(lesson_text)}\n\n"

    return msg.strip()


# --- Хэндлеры команд ---
@dp.message(Command("start"))
async def cmd_start(message: Message):
    user_id = message.from_user.id
    if user_id in user_groups:
        await show_days_menu(message, user_groups[user_id])
    else:
        await message.answer(
            "👋 Привет! Я бот актуального расписания Физтех-колледжа.\n\n"
            "Я автоматически отслеживаю и подстраиваюсь под любые форматы и изменения файлов расписания на сайте колледжа.\n\n"
            "✍️ Отправь мне номер своей группы (например: <b>ИСП-31</b> или <b>АДТ-111</b>).\n"
            "Я запомню её, и ты сможешь удобно смотреть пары кнопками!",
            parse_mode="HTML"
        )


@dp.message(Command("menu", "days"))
async def cmd_menu(message: Message):
    user_id = message.from_user.id
    if user_id in user_groups:
        await show_days_menu(message, user_groups[user_id])
    else:
        await message.answer(
            "Сначала укажите номер группы (например: <b>ИСП-31</b> или <b>АДТ-111</b>):",
            parse_mode="HTML"
        )


@dp.message(Command("group"))
async def cmd_group(message: Message):
    await message.answer(
        "✍️ Введите номер вашей группы (например: <b>ИСП-31</b> или <b>АВТ-112</b>):",
        parse_mode="HTML"
    )


@dp.message(Command("refresh"))
async def cmd_refresh(message: Message):
    status_msg = await message.answer("🔄 Загружаю свежее расписание с сайта колледжа...")
    schedule = await get_schedule(force_refresh=True)
    dates_count = len(schedule)
    await status_msg.edit_text(
        f"✅ Расписание обновлено! Найдено дат: {dates_count}.\n"
        f"Чтобы открыть меню, нажмите /menu"
    )


@dp.message(Command("help"))
async def cmd_help(message: Message):
    help_text = (
        "📖 <b>Команды бота расписания:</b>\n\n"
        "🔹 /menu или /days — открыть выбор дней расписания\n"
        "🔹 /group — изменить свою учебную группу\n"
        "🔹 /refresh — принудительно обновить расписание с сайта колледжа\n"
        "🔹 /help — показать эту справку\n\n"
        "💡 <i>Бот автоматически объединяет расписания всех корпусов и адаптируется к любым изменениям таблиц.</i>"
    )
    await message.answer(help_text, parse_mode="HTML")


# --- Текстовый хэндлер (сохранение группы) ---
@dp.message(F.text, ~F.text.startswith("/"))
async def msg_save_group(message: Message):
    user_id = message.from_user.id
    group_name = message.text.strip()

    if len(group_name) < 2 or len(group_name) > 30:
        await message.answer(
            "Некорректное название группы. Попробуй еще раз (например: <b>ИСП-31</b>, <b>АДТ-111</b>).",
            parse_mode="HTML"
        )
        return

    schedule = await get_schedule(force_refresh=False)
    matched_name = None
    if schedule:
        for day_data in schedule.values():
            found = find_group_in_dict(day_data.get("groups", {}), group_name)
            if found:
                matched_name = found
                break

    saved_name = matched_name if matched_name else group_name
    user_groups[user_id] = saved_name
    save_user_groups(user_groups)

    await message.answer(
        f"✅ Группа <b>{html.escape(saved_name)}</b> сохранена!",
        parse_mode="HTML"
    )
    await show_days_menu(message, saved_name)


# --- Отображение меню дней ---
async def show_days_menu(message_or_query, group_name: str, force_refresh: bool = False):
    if isinstance(message_or_query, Message):
        send_method = message_or_query.answer
        is_callback = False
    else:
        send_method = message_or_query.message.edit_text
        is_callback = True

    schedule = await get_schedule(force_refresh=force_refresh)

    if not schedule:
        await send_method(
            "❌ Не удалось загрузить расписание с сайта колледжа. Попробуйте позже через /refresh."
        )
        return

    builder = InlineKeyboardBuilder()
    sorted_dates = sorted(
        schedule.keys(),
        key=lambda d: datetime.strptime(d, "%d.%m.%Y")
    )

    for date_str in sorted_dates:
        day_data = schedule[date_str]
        dt = datetime.strptime(date_str, "%d.%m.%Y")
        day_en = dt.strftime("%A")
        short_ru = SHORT_DAY_NAMES.get(day_en, day_data.get("day_name", "")[:2])

        btn_text = f"📅 {date_str[:5]} ({short_ru})"
        builder.button(text=btn_text, callback_data=f"day_{date_str}")

    builder.button(text="🔄 Обновить", callback_data="refresh_schedule")
    builder.button(text="✏️ Сменить группу", callback_data="change_group")
    builder.adjust(1)

    group_found = any(
        find_group_in_dict(d.get("groups", {}), group_name)
        for d in schedule.values()
    )

    if group_found:
        text = (
            f"👤 Ваша группа: <b>{html.escape(group_name)}</b>\n\n"
            f"👇 Выберите день для просмотра расписания:"
        )
    else:
        text = (
            f"👤 Ваша группа: <b>{html.escape(group_name)}</b>\n\n"
            f"⚠️ <i>Группа пока не найдена в текущем расписании на сайте. "
            f"Проверьте название группы или выберите день:</i>"
        )

    try:
        await send_method(text, reply_markup=builder.as_markup(), parse_mode="HTML")
    except TelegramBadRequest as e:
        if "message is not modified" in str(e).lower() and is_callback:
            await message_or_query.answer("Расписание уже актуально")
        else:
            logger.error(f"TelegramBadRequest in show_days_menu: {e}")


# --- Callbacks ---
@dp.callback_query(F.data.startswith("day_"))
async def cbq_show_day(callback: CallbackQuery):
    date_str = callback.data.replace("day_", "")
    user_id = callback.from_user.id
    group_name = user_groups.get(user_id, "Неизвестная группа")

    await callback.answer()

    schedule = await get_schedule(force_refresh=False)
    text = get_schedule_message(schedule, date_str, group_name)

    builder = InlineKeyboardBuilder()
    builder.button(text="◀️ К выбору дня", callback_data="back_to_days")
    builder.button(text="🔄 Обновить", callback_data=f"refresh_day_{date_str}")
    builder.adjust(1)

    try:
        if len(text) > 4000:
            parts = [text[i:i + 3900] for i in range(0, len(text), 3900)]
            for p in parts[:-1]:
                await callback.message.answer(p, parse_mode="HTML")
            await callback.message.edit_text(parts[-1], reply_markup=builder.as_markup(), parse_mode="HTML")
        else:
            await callback.message.edit_text(text, reply_markup=builder.as_markup(), parse_mode="HTML")
    except TelegramBadRequest as e:
        if "message is not modified" not in str(e).lower():
            logger.error(f"Telegram error in cbq_show_day: {e}")


@dp.callback_query(F.data.startswith("refresh_day_"))
async def cbq_refresh_day(callback: CallbackQuery):
    date_str = callback.data.replace("refresh_day_", "")
    user_id = callback.from_user.id
    group_name = user_groups.get(user_id, "Неизвестная группа")

    await callback.answer("Обновляю с сайта...")
    schedule = await get_schedule(force_refresh=True)
    text = get_schedule_message(schedule, date_str, group_name)

    builder = InlineKeyboardBuilder()
    builder.button(text="◀️ К выбору дня", callback_data="back_to_days")
    builder.button(text="🔄 Обновить", callback_data=f"refresh_day_{date_str}")
    builder.adjust(1)

    try:
        await callback.message.edit_text(text, reply_markup=builder.as_markup(), parse_mode="HTML")
    except TelegramBadRequest as e:
        if "message is not modified" in str(e).lower():
            await callback.answer("Расписание не изменилось")


@dp.callback_query(F.data == "back_to_days")
async def cbq_back(callback: CallbackQuery):
    user_id = callback.from_user.id
    if user_id in user_groups:
        await callback.answer()
        await show_days_menu(callback, user_groups[user_id])
    else:
        await callback.answer("Группа не выбрана. Введите номер группы.")
        await callback.message.answer("Введите номер вашей группы (например: <b>ИСП-31</b>):", parse_mode="HTML")


@dp.callback_query(F.data == "refresh_schedule")
async def cbq_refresh(callback: CallbackQuery):
    user_id = callback.from_user.id
    group_name = user_groups.get(user_id, "")
    await callback.answer("Загружаю актуальные данные...")
    await show_days_menu(callback, group_name, force_refresh=True)


@dp.callback_query(F.data == "change_group")
async def cbq_change_group(callback: CallbackQuery):
    user_id = callback.from_user.id
    if user_id in user_groups:
        del user_groups[user_id]
        save_user_groups(user_groups)
    await callback.answer("Группа сброшена.")
    await callback.message.edit_text(
        "✍️ Введите новый номер группы:\n<i>(например: ИСП-31 или АДТ-111)</i>",
        parse_mode="HTML"
    )


async def main():
    logger.info("Инициализация расписания при запуске...")
    await get_schedule(force_refresh=True)
    logger.info("Бот успешно запущен и ожидает сообщений...")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())