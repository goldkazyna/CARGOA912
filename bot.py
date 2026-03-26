import logging
import os
import sqlite3
from io import BytesIO
from openpyxl import Workbook
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ConversationHandler,
    filters,
)

# --- Config ---
BOT_TOKEN = "8791855388:AAEbwTATC13rwM1LKvNjjhbHB69yg1_PQZs"  # Вставь свой токен
ADMIN_IDS = [381314146, 634620925]  # Telegram ID админов

# --- Logging ---
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# --- States ---
NAME, PHONE, CITY, ADDRESS = range(4)

# --- Database ---
def init_db():
    conn = sqlite3.connect("clients.db")
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS clients (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            telegram_id INTEGER UNIQUE,
            name TEXT,
            phone TEXT,
            city TEXT,
            address TEXT,
            client_code TEXT
        )
    """)
    conn.commit()
    conn.close()


def get_next_client_number():
    conn = sqlite3.connect("clients.db")
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM clients")
    count = c.fetchone()[0]
    conn.close()
    return count + 1


def save_client(telegram_id, name, phone, city, address):
    conn = sqlite3.connect("clients.db")
    c = conn.cursor()

    c.execute("SELECT id FROM clients WHERE telegram_id = ?", (telegram_id,))
    existing = c.fetchone()

    number = get_next_client_number() if not existing else None
    client_code = f"А912–ALA1-{number:06d}" if number else None

    if existing:
        c.execute(
            "UPDATE clients SET name=?, phone=?, city=?, address=? WHERE telegram_id=?",
            (name, phone, city, address, telegram_id),
        )
        c.execute("SELECT client_code FROM clients WHERE telegram_id=?", (telegram_id,))
        client_code = c.fetchone()[0]
    else:
        c.execute(
            "INSERT INTO clients (telegram_id, name, phone, city, address, client_code) VALUES (?, ?, ?, ?, ?, ?)",
            (telegram_id, name, phone, city, address, client_code),
        )

    conn.commit()
    conn.close()
    return client_code


# --- Links ---
LINKS_KEYBOARD = InlineKeyboardMarkup([
    [InlineKeyboardButton("📞 Написать менеджеру", url="https://clck.ru/3SixDm")],
    [InlineKeyboardButton("💬 Группа в Telegram", url="https://t.me/cargoA912")],
])


def build_final_message(name, client_code, city, address):
    return (
        f"Спасибо, {name}. Теперь ты стал нашим клиентом 😁\n\n"
        f"Ваш личный код:\n"
        f"<b>{client_code}</b>\n\n"
        f"Город: {city}\n"
        f"Адрес пункта выдачи: {address}\n\n"
        f"Адрес для заполнения склада в Китае ⬇️\n\n"
        f"1) {client_code}\n"
        f"2) 18618148777\n"
        f"3) 浙江省 金华市 义乌市\n"
        f"4) 浙江省金华市义乌市陶界岭18幢A186  吴立斌 ( {client_code} )\n\n"
        f"ЕСЛИ СОМНЕВАЕТЕСЬ ЧТО НЕПРАВИЛЬНО ЗАПОЛНИЛИ ИЛИ НЕ МОЖЕТЕ АДРЕС, "
        f"НАПИШИТЕ НАШЕМУ МЕНЕДЖЕРУ‼️‼️‼️\n\n"
        f"Адрес склада ⬇️\n"
        f"Жарокова 12"
    )


# --- Handlers ---
async def start(update: Update, context):
    conn = sqlite3.connect("clients.db")
    c = conn.cursor()
    c.execute("SELECT name, client_code, city, address FROM clients WHERE telegram_id = ?", (update.effective_user.id,))
    row = c.fetchone()
    conn.close()

    if row:
        name, client_code, city, address = row
        message = build_final_message(name, client_code, city, address)
        await update.message.reply_text(message, parse_mode="HTML", reply_markup=LINKS_KEYBOARD)
        return ConversationHandler.END

    context.user_data.clear()
    await update.message.reply_text(
        "Добро пожаловать! 👋\n\nНапишите своё имя в чат:"
    )
    return NAME


async def get_name(update: Update, context):
    context.user_data["name"] = update.message.text
    await update.message.reply_text("Напишите Ваш номер телефона:")
    return PHONE


async def get_phone(update: Update, context):
    context.user_data["phone"] = update.message.text
    keyboard = [
        [InlineKeyboardButton("Алматы", callback_data="city_Алматы")],
    ]
    await update.message.reply_text(
        "Выберите город получения заказов:",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    return CITY


async def get_city(update: Update, context):
    query = update.callback_query
    await query.answer()
    city = query.data.replace("city_", "")
    context.user_data["city"] = city

    keyboard = [
        [InlineKeyboardButton("Жарокова 12", callback_data="addr_Жарокова 12")],
    ]
    await query.edit_message_text(
        f"Город: {city} ✅\n\nВыберите адрес пункта выдачи:",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    return ADDRESS


async def get_address(update: Update, context):
    query = update.callback_query
    await query.answer()
    address = query.data.replace("addr_", "")
    context.user_data["address"] = address

    await query.edit_message_text(f"Адрес: {address} ✅")

    client_code = save_client(
        telegram_id=update.effective_user.id,
        name=context.user_data["name"],
        phone=context.user_data["phone"],
        city=context.user_data["city"],
        address=context.user_data["address"],
    )

    name = context.user_data["name"]
    city = context.user_data["city"]
    address = context.user_data["address"]
    message = build_final_message(name, client_code, city, address)

    await query.message.reply_text(message, parse_mode="HTML", reply_markup=LINKS_KEYBOARD)
    return ConversationHandler.END


async def cancel(update: Update, context):
    await update.message.reply_text("Отменено.")
    return ConversationHandler.END


# --- Admin ---
async def admin(update: Update, context):
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("У вас нет доступа.")
        return

    keyboard = [[InlineKeyboardButton("📊 Выгрузить в Excel", callback_data="admin_export")]]
    await update.message.reply_text(
        "Админ-панель:",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )


async def admin_export(update: Update, context):
    query = update.callback_query
    if query.from_user.id not in ADMIN_IDS:
        await query.answer("Нет доступа", show_alert=True)
        return

    await query.answer()

    conn = sqlite3.connect("clients.db")
    c = conn.cursor()
    c.execute("SELECT client_code, name, phone, city, address, telegram_id FROM clients ORDER BY id")
    rows = c.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Клиенты"
    ws.append(["Код клиента", "Имя", "Телефон", "Город", "Адрес", "Telegram ID"])

    for row in rows:
        ws.append(list(row))

    # Автоширина колонок
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 2

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    await query.message.reply_document(
        document=buffer,
        filename="clients.xlsx",
        caption=f"Всего клиентов: {len(rows)}",
    )


# --- Main ---
def main():
    init_db()

    app = Application.builder().token(BOT_TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_name)],
            PHONE: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_phone)],
            CITY: [CallbackQueryHandler(get_city, pattern=r"^city_")],
            ADDRESS: [CallbackQueryHandler(get_address, pattern=r"^addr_")],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(conv_handler)
    app.add_handler(CommandHandler("admin", admin))
    app.add_handler(CallbackQueryHandler(admin_export, pattern=r"^admin_export$"))

    logger.info("Бот запущен!")
    app.run_polling()


if __name__ == "__main__":
    main()
