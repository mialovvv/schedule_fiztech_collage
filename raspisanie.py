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

logging.basicConfig(level=logging.INFO)

BOT_TOKEN = "bot_token"

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

user_groups = {}


def get_latest_schedule_url() -> str:
    url = "https://phystech.pro/%D1%81%D1%82%D1%83%D0%B4%D0%B5%D0%BD%D1%82%D1%83/%D1%80%D0%B0%D1%81%D0%BF%D0%B8%D1%81%D0%B0%D0%BD%D0%B8%D0%B5/"
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        html = urllib.request.urlopen(req).read()
        soup = BeautifulSoup(html, 'html.parser')
        links = soup.find_all('a', href=True)

        xlsx_links = []
        for link in links:
            href = link['href']
            if '.xlsx' in href:
                dates = re.findall(r'(\d{2}\.\d{2})', href)
                if len(dates) >= 2:
                    date_str = dates[0]
                    try:
                        dt = datetime.strptime(f"{date_str}.{datetime.now().year}", "%d.%m.%Y")
                        xlsx_links.append((dt, href))
                    except ValueError:
                        pass

        if not xlsx_links:
            return "https://phystech.pro/wp-content/uploads/2026/09/%D0%A1%D0%9F-%E2%84%961-%D1%81-09.09-%D0%BF%D0%BE-11.09.xlsx"

        xlsx_links.sort(key=lambda x: x[0], reverse=True)
        latest_href = xlsx_links[0][1]
        if latest_href.startswith('/'):
            latest_href = "https://phystech.pro" + latest_href
        return urllib.parse.quote(latest_href, safe=':/')
    except Exception as e:
        logging.error(f"Ошибка поиска URL: {e}")
        return "https://phystech.pro/wp-content/uploads/2026/09/%D0%A1%D0%9F-%E2%84%961-%D1%81-09.09-%D0%BF%D0%BE-11.09.xlsx"


def parse_schedule(url: str) -> dict:
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    response = urllib.request.urlopen(req)
    data = response.read()

    df = pd.read_excel(BytesIO(data), sheet_name=None, header=None)
    sheet_df = list(df.values())[0]

    filename = url.split('/')[-1]
    dates = re.findall(r'(\d{2}\.\d{2})', filename)
    if len(dates) >= 2:
        start_date = datetime.strptime(f"{dates[0]}.{datetime.now().year}", "%d.%m.%Y")
    else:
        start_date = datetime.now()

    header_rows = []
    for i in range(len(sheet_df)):
        col2 = sheet_df.iloc[i, 2]
        if col2 == 'Пара' or (pd.isna(col2) and str(col2) == 'nan'):
            row_vals = sheet_df.iloc[i, 3:].dropna().astype(str)
            if any('/' in val or val.strip().endswith('/26') or val.strip() == 'Пара' for val in row_vals):
                header_rows.append(i)

    if 0 not in header_rows:
        row0_vals = sheet_df.iloc[0, 3:].dropna().astype(str)
        if any('/' in val or val.strip().endswith('/26') for val in row0_vals):
            header_rows.insert(0, 0)

    schedule_by_date = {}
    for idx, header_idx in enumerate(header_rows):
        current_date = start_date + timedelta(days=idx)
        date_str = current_date.strftime("%d.%m.%Y")
        day_name = current_date.strftime("%A")

        if idx + 1 < len(header_rows):
            end_idx = header_rows[idx + 1]
        else:
            end_idx = len(sheet_df)

        header_row = sheet_df.iloc[header_idx]
        group_cols = {}
        for col_idx in range(3, len(header_row)):
            group_name = header_row.iloc[col_idx]
            if pd.notna(group_name) and str(group_name).strip():
                group_cols[str(group_name).strip()] = col_idx

        day_schedule = {}
        for row_idx in range(header_idx + 1, end_idx):
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

        schedule_by_date[date_str] = {
            "day_name": day_name,
            "groups": day_schedule
        }
    return schedule_by_date


def find_group_in_dict(groups_dict: dict, target_group: str) -> str:
    target_clean = target_group.strip().upper().replace(" ", "").replace("-", "")
    for g in groups_dict.keys():
        g_clean = g.upper().replace(" ", "").replace("-", "")
        if target_clean in g_clean or g_clean in target_clean:
            return g
    return None


def get_schedule_message(schedule_data: dict, date_str: str, group_name: str) -> str:
    if date_str not in schedule_data:
        return "Ошибка: данные за этот день не найдены."

    day_data = schedule_data[date_str]
    matched_group = find_group_in_dict(day_data["groups"], group_name)

    if not matched_group:
        available = ", ".join(list(day_data["groups"].keys())[:8]) + "..."
        return f"Группа '{group_name}' не найдена на эту дату.\nДоступные группы: {available}"

    lessons = day_data["groups"][matched_group]
    if not lessons:
        return f"На этот день у группы {matched_group} пар нет! Отдыхай."

    day_ru = {"Monday": "Понедельник", "Tuesday": "Вторник", "Wednesday": "Среда",
              "Thursday": "Четверг", "Friday": "Пятница", "Saturday": "Суббота", "Sunday": "Воскресенье"}
    day_name_ru = day_ru.get(day_data["day_name"], day_data["day_name"])

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
            "Чтобы начать, отправь мне номер своей группы (например: ИСП-31 или АДТ-111).\n"
            "Я запомню её, и дальше ты сможешь выбирать дни кнопками!"
        )


@dp.message(F.text)
async def msg_save_group(message: Message):
    user_id = message.from_user.id
    group_name = message.text.strip()

    if len(group_name) < 3:
        await message.answer("Название группы слишком короткое. Попробуй еще раз (например: ИСП-31).")
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
        target = message_or_query
        send_method = target.answer
    else:
        target = message_or_query.message
        send_method = target.edit_text

    loading_msg = await send_method("Загружаю актуальное расписание с сайта колледжа...")

    try:
        url = await asyncio.to_thread(get_latest_schedule_url)
        schedule = await asyncio.to_thread(parse_schedule, url)

        builder = InlineKeyboardBuilder()
        available_dates = []

        sorted_dates = sorted(schedule.keys(), key=lambda d: datetime.strptime(d, "%d.%m.%Y"))

        for date_str in sorted_dates:
            day_data = schedule[date_str]
            day_name = day_data["day_name"]
            day_ru = {"Monday": "Пн", "Tuesday": "Вт", "Wednesday": "Ср",
                      "Thursday": "Чт", "Friday": "Пт", "Saturday": "Сб", "Sunday": "Вс"}

            btn_text = f"{date_str[:5]} ({day_ru.get(day_name, day_name[:2])})"
            builder.button(text=btn_text, callback_data=f"day_{date_str}")
            available_dates.append(date_str)

        builder.button(text="Изменить группу", callback_data="change_group")
        builder.adjust(1)

        group_found_anywhere = any(find_group_in_dict(d["groups"], group_name) for d in schedule.values())

        if not group_found_anywhere and available_dates:
            sample_date = available_dates[0]
            sample_msg = get_schedule_message(schedule, sample_date, group_name)
            await loading_msg.edit_text(sample_msg + "\n\nВыберите другой день или измените группу:", reply_markup=builder.as_markup())
        else:
            await loading_msg.edit_text(f"Твоя группа: {group_name}\nВыбери день:", reply_markup=builder.as_markup())

    except Exception as e:
        logging.error(e)
        await loading_msg.edit_text("Не удалось загрузить расписание. Попробуй позже.")


@dp.callback_query(F.data.startswith("day_"))
async def cbq_show_day(callback: CallbackQuery):
    date_str = callback.data.replace("day_", "")
    user_id = callback.from_user.id
    group_name = user_groups.get(user_id, "Неизвестная группа")

    await callback.answer("Загружаю...")

    try:
        url = await asyncio.to_thread(get_latest_schedule_url)
        schedule = await asyncio.to_thread(parse_schedule, url)

        text = get_schedule_message(schedule, date_str, group_name)

        builder = InlineKeyboardBuilder()
        builder.button(text="К выбору дня", callback_data="back_to_days")

        await callback.message.edit_text(text, reply_markup=builder.as_markup())
    except Exception as e:
        logging.error(e)
        await callback.message.edit_text("Ошибка при получении данных.")


@dp.callback_query(F.data == "back_to_days")
async def cbq_back(callback: CallbackQuery):
    user_id = callback.from_user.id
    if user_id in user_groups:
        await callback.answer()
        await show_days_menu(callback, user_groups[user_id])
    else:
        await callback.answer("Группа не найдена, введите /start")


@dp.callback_query(F.data == "change_group")
async def cbq_change_group(callback: CallbackQuery):
    user_id = callback.from_user.id
    if user_id in user_groups:
        del user_groups[user_id]
    await callback.answer("Группа сброшена. Введите новую.")
    await callback.message.edit_text("Введите новый номер группы:\n(например: ИСП-32)")


async def main():
    print("Бот успешно запущен и ждет сообщений...")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())