
import os
import re
import logging
import asyncio
import urllib.request
import urllib.parse
from datetime import datetime, timedelta
from io import BytesIO

import pandas as pd
from bs4 import BeautifulSoup
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.types import Message, CallbackQuery
from aiogram.utils.keyboard import InlineKeyboardBuilder
from dotenv import load_dotenv

logging.basicConfig(level=logging.INFO)
load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN", "YOUR_BOT_TOKEN_HERE")

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
user_groups = {}

DAY_RU = {"Monday": "Пн", "Tuesday": "Вт", "Wednesday": "Ср",
          "Thursday": "Чт", "Friday": "Пт", "Saturday": "Сб", "Sunday": "Вс"}
DAY_RU_FULL = {"Monday": "Понедельник", "Tuesday": "Вторник", "Wednesday": "Среда",
               "Thursday": "Четверг", "Friday": "Пятница", "Saturday": "Суббота", "Sunday": "Воскресенье"}

def encode_url(url: str) -> str:
    try:
        url.encode('ascii')
        return url
    except UnicodeEncodeError:
        parsed = urllib.parse.urlparse(url)
        path_parts = parsed.path.split('/')
        encoded_path = '/'.join(urllib.parse.quote(part, safe='') for part in path_parts)
        return urllib.parse.urlunparse(parsed._replace(path=encoded_path))

def get_all_schedule_urls() -> list:
    url = "https://phystech.pro/%D1%81%D1%82%D1%83%D0%B4%D0%B5%D0%BD%D1%82%D1%83/%D1%80%D0%B0%D1%81%D0%BF%D0%B8%D1%81%D0%B0%D0%BD%D0%B8%D0%B5/"
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        html = urllib.request.urlopen(req).read()
        soup = BeautifulSoup(html, 'html.parser')
        links = soup.find_all('a', href=True)
        
        xlsx_links = []
        for link in links:
            href = link['href']
            if '.xlsx' not in href:
                continue
                
            filename = href.split('/')[-1]
            logging.info(f"Найден файл: {filename}")
            
            # Формат 1: 26-09-14-... (ГГ-ММ-ДД)
            match1 = re.match(r'^(\d{2})-(\d{2})-(\d{2})-', filename)
            if match1:
                yy, mm, dd = match1.groups()
                year = 2000 + int(yy)
                try:
                    dt = datetime(year, int(mm), int(dd))
                    xlsx_links.append((dt, href))
                    continue
                except:
                    pass
            
            # Формат 1б: 2026-09-14-... (ГГГГ-ММ-ДД)
            match1b = re.match(r'^(\d{4})-(\d{2})-(\d{2})-', filename)
            if match1b:
                yyyy, mm, dd = match1b.groups()
                try:
                    dt = datetime(int(yyyy), int(mm), int(dd))
                    xlsx_links.append((dt, href))
                    continue
                except:
                    pass
            
            # Формат 2: с-09.09-по-11.09
            match2 = re.search(r'с[- ]?(\d{2}\.\d{2})[- ]?по[- ]?(\d{2}\.\d{2})', filename)
            if match2:
                date_str = match2.group(1)
                try:
                    dt = datetime.strptime(f"{date_str}.{datetime.now().year}", "%d.%m.%Y")
                    xlsx_links.append((dt, href))
                    continue
                except:
                    pass
            
            # Формат 3: на 14.09.2026
            match3 = re.search(r'на[ ]?(\d{2}\.\d{2})\.?(\d{4})?', filename)
            if match3:
                date_str = match3.group(1)
                year = match3.group(2) if match3.group(2) else str(datetime.now().year)
                try:
                    dt = datetime.strptime(f"{date_str}.{year}", "%d.%m.%Y")
                    xlsx_links.append((dt, href))
                    continue
                except:
                    pass
        
        xlsx_links.sort(key=lambda x: x[0], reverse=True)
        logging.info(f"Всего найдено файлов расписания: {len(xlsx_links)}")
        for dt, href in xlsx_links:
            logging.info(f"  {dt.strftime('%d.%m.%Y')}: {href}")
        return xlsx_links
        
    except Exception as e:
        logging.error(f"Ошибка поиска URL: {e}")
        return []

def parse_single_file(url: str) -> dict:
    safe_url = encode_url(url)
    try:
        req = urllib.request.Request(safe_url, headers={'User-Agent': 'Mozilla/5.0'})
        response = urllib.request.urlopen(req)
        data = response.read()
        df = pd.read_excel(BytesIO(data), sheet_name=None, header=None)
        sheet_df = list(df.values())[0]
    except Exception as e:
        logging.error(f"Ошибка чтения {safe_url}: {e}")
        return {}
    
    first_cell = str(sheet_df.iloc[0, 0]).strip()
    match_date = re.match(r'^(\d{4})-(\d{2})-(\d{2})$', first_cell)
    if match_date:
        y, m, d = match_date.groups()
        file_date = datetime(int(y), int(m), int(d))
        header_row_idx = 1
        data_start_idx = 2
    else:
        filename = url.split('/')[-1]
        dates = re.findall(r'(\d{2}\.\d{2})', filename)
        if dates:
            file_date = datetime.strptime(f"{dates[0]}.{datetime.now().year}", "%d.%m.%Y")
        else:
            file_date = datetime.now()
        header_row_idx = 0
        data_start_idx = 1
    
    date_str = file_date.strftime("%d.%m.%Y")
    day_name = file_date.strftime("%A")
    
    header_row = sheet_df.iloc[header_row_idx]
    group_cols = {}
    for col_idx in range(3, len(header_row)):
        group_name = header_row.iloc[col_idx]
        if pd.notna(group_name) and str(group_name).strip():
            group_cols[str(group_name).strip()] = col_idx
    
    day_schedule = {}
    for row_idx in range(data_start_idx, len(sheet_df)):
        row = sheet_df.iloc[row_idx]
        lesson_num = row.iloc[2]
        
        if pd.isna(lesson_num) or str(lesson_num).strip() == 'Перерыв' or str(lesson_num).strip() == '':
            continue
        
        time_str = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        
        for group_name, col_idx in group_cols.items():
            lesson_info = row.iloc[col_idx]
            if pd.notna(lesson_info) and str(lesson_info).strip():
                if group_name not in day_schedule:
                    day_schedule[group_name] = []
                day_schedule[group_name].append({
                    "pair": str(lesson_num).strip(),
                    "time": time_str,
                    "lesson": str(lesson_info).strip()
                })
    
    return {date_str: {"day_name": day_name, "groups": day_schedule}}

def parse_all_schedules() -> dict:
    all_urls = get_all_schedule_urls()
    if not all_urls:
        return {}
    
    combined = {}
    for dt, url in all_urls:
        file_data = parse_single_file(url)
        for date_str, data in file_data.items():
            if date_str not in combined:
                combined[date_str] = data
            else:
                for group, lessons in data["groups"].items():
                    if group not in combined[date_str]["groups"]:
                        combined[date_str]["groups"][group] = lessons
                    else:
                        combined[date_str]["groups"][group].extend(lessons)
    
    return combined

def find_group_in_dict(groups_dict: dict, target_group: str) -> str:
    target_clean = target_group.strip().upper().replace(" ", "").replace("-", "")
    for g in groups_dict.keys():
        g_clean = g.upper().replace(" ", "").replace("-", "").replace("/26", "")
        if target_clean in g_clean or g_clean in target_clean:
            return g
    return None

def get_schedule_message(schedule_data: dict, date_str: str, group_name: str) -> str:
    if date_str not in schedule_data:
        return f"На {date_str} расписание не найдено."
    day_data = schedule_data[date_str]
    matched_group = find_group_in_dict(day_data["groups"], group_name)
    if not matched_group:
        available = ", ".join(list(day_data["groups"].keys())[:10])
        return f"Группа '{group_name}' не найдена на {date_str}.\nДоступные группы: {available}"
    lessons = day_data["groups"][matched_group]
    if not lessons:
        return f"На {date_str} у группы {matched_group} пар нет!"
    day_name_ru = DAY_RU_FULL.get(day_data["day_name"], day_data["day_name"])
    response = f"Расписание: {matched_group}\n{date_str} ({day_name_ru})\n\n"
    for lesson in lessons:
        response += f"{lesson['pair']} пара ({lesson['time']}):\n{lesson['lesson']}\n\n"
    return response.strip()

@dp.message(Command("start"))
async def cmd_start(message: Message):
    user_id = message.from_user.id
    if user_id in user_groups:
        await show_days_menu(message, user_groups[user_id])
    else:
        await message.answer(
            "Привет! Я бот расписания Физтех-колледжа.\n\n"
            "Отправь мне номер своей группы (например: ИСП-31 или АДТ-111).\n"
            "Я запомню её, и дальше ты сможешь выбирать дни кнопками!"
        )

@dp.message(F.text)
async def msg_save_group(message: Message):
    user_id = message.from_user.id
    group_name = message.text.strip()
    if len(group_name) < 3:
        await message.answer("Название группы слишком короткое. Попробуй ещё раз (например: ИСП-31).")
        return
    user_groups[user_id] = group_name
    await message.answer(f"Группа {group_name} сохранена!")
    await show_days_menu(message, group_name)

@dp.message(Command("menu", "days"))
async def cmd_menu(message: Message):
    user_id = message.from_user.id
    if user_id in user_groups:
        await show_days_menu(message, user_groups[user_id])
    else:
        await cmd_start(message)

async def show_days_menu(message_or_query, group_name: str):
    if isinstance(message_or_query, Message):
        send_method = message_or_query.answer
    else:
        send_method = message_or_query.message.edit_text

    loading_msg = await send_method("Загружаю расписание с сайта колледжа...")

    try:
        schedule = await asyncio.to_thread(parse_all_schedules)
        
        if not schedule:
            await loading_msg.edit_text("Не удалось найти расписание на сайте.")
            return

        builder = InlineKeyboardBuilder()
        sorted_dates = sorted(schedule.keys(), key=lambda d: datetime.strptime(d, "%d.%m.%Y"))

        for date_str in sorted_dates:
            day_data = schedule[date_str]
            day_name = day_data["day_name"]
            btn_text = f"{date_str[:5]} ({DAY_RU.get(day_name, day_name[:2])})"
            builder.button(text=btn_text, callback_data=f"day_{date_str}")

        builder.button(text="Изменить группу", callback_data="change_group")
        builder.adjust(1)

        group_found = any(find_group_in_dict(d["groups"], group_name) for d in schedule.values())
        if not group_found:
            sample = sorted_dates[0]
            sample_msg = get_schedule_message(schedule, sample, group_name)
            await loading_msg.edit_text(sample_msg + "\n\nВыберите другой день или измените группу:", reply_markup=builder.as_markup())
        else:
            await loading_msg.edit_text(f"Твоя группа: {group_name}\nВыбери день:", reply_markup=builder.as_markup())

    except Exception as e:
        logging.error(e)
        await loading_msg.edit_text(f"Не удалось загрузить расписание. Ошибка: {e}")

@dp.callback_query(F.data.startswith("day_"))
async def cbq_show_day(callback: CallbackQuery):
    date_str = callback.data.replace("day_", "")
    user_id = callback.from_user.id
    group_name = user_groups.get(user_id, "Неизвестная группа")
    await callback.answer("Загружаю...")
    try:
        schedule = await asyncio.to_thread(parse_all_schedules)
        text = get_schedule_message(schedule, date_str, group_name)
        builder = InlineKeyboardBuilder()
        builder.button(text="К выбору дня", callback_data="back_to_days")
        await callback.message.edit_text(text, reply_markup=builder.as_markup())
    except Exception as e:
        logging.error(e)
        await callback.message.edit_text(f"Ошибка: {e}")

@dp.callback_query(F.data == "back_to_days")
async def cbq_back(callback: CallbackQuery):
    user_id = callback.from_user.id
    if user_id in user_groups:
        await callback.answer()
        await show_days_menu(callback, user_groups[user_id])

@dp.callback_query(F.data == "change_group")
async def cbq_change_group(callback: CallbackQuery):
    user_id = callback.from_user.id
    if user_id in user_groups:
        del user_groups[user_id]
    await callback.answer("Группа сброшена. Введите новую.")
    await callback.message.edit_text("Введите новый номер группы (например: ИСП-32)")

async def main():
    print("Бот запущен...")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
