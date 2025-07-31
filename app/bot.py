from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes, MessageHandler, filters
from my_token import TOKEN
from xlsx_tools.tables import fill_template
import os
from decimal import Decimal
user_data = {}  # user_id: {'inc_category': [], 'inc_sum': [], 'exp_category': [], 'exp_sum': []}

# Стартовое сообщение
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message:
        await update.message.reply_text(
            "Привет! Я бот учёта доходов и расходов.\n"
            "Отправьте мне:\n"
            "Доход: <категория> <сумма>\n"
            "Расход: <категория> <сумма>\n"
            "Пример:\n"
            "Доход: Зарплата 100000\n" 
            "Доход: Дивиденды 10000\n"
            "Расход: Продукты 12350\n"
            "А чтобы получить отчёт, напишите /doit"
        )

# Справка
async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message:
        await update.message.reply_text(
            "Это справка по боту.\n"
            "Формат сообщений:\n"
            "Доход: <категория> <сумма>\n"
            "Расход: <категория> <сумма>\n"
            "Для отчёта: /doit"
        )

# Обработчик текстовых сообщений
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message and update.effective_user:
        text = update.message.text
        if text is not None:
            text = text.strip()
            user_id = update.effective_user.id
            user_data.setdefault(user_id, {
                'inc_category': [], 'inc_sum': [],
                'exp_category': [], 'exp_sum': []
            })
            if text.lower().startswith("доход:"):
                try:
                    _, rest = text.split(":", 1)
                    parts = rest.strip().rsplit(" ", 1)
                    category = parts[0].strip()
                    summ = Decimal(parts[1].replace(",", "."))
                except Exception:
                    await update.message.reply_text("Формат: Доход: <категория> <сумма>")
                    return
                user_data[user_id]['inc_category'].append(category)
                user_data[user_id]['inc_sum'].append(summ)
                await update.message.reply_text(f"Доход '{category}' на сумму {summ} записан!")
            elif text.lower().startswith("расход:"):
                try:
                    _, rest = text.split(":", 1)
                    parts = rest.strip().rsplit(" ", 1)
                    category = parts[0].strip()
                    summ = Decimal(parts[1].replace(",", "."))
                except Exception:
                    await update.message.reply_text("Формат: Расход: <категория> <сумма>")
                    return
                user_data[user_id]['exp_category'].append(category)
                user_data[user_id]['exp_sum'].append(summ)
                await update.message.reply_text(f"Расход '{category}' на сумму {summ} записан!")
            else:
                await update.message.reply_text("Пожалуйста, используйте формат: Доход: <категория> <сумма> или Расход: <категория> <сумма>")

# Команда /doit — отправка Excel-отчёта
async def doit(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message and update.effective_user:
        user_id = update.effective_user.id
        data = user_data.get(user_id)
        if not data or (not data['inc_category'] and not data['exp_category']):
            await update.message.reply_text("Нет данных для отчёта.")
            return
        file_path = f"income_{user_id}.xlsx"
        template_path = "app/files/table_test.xlsx"
        fill_template(
            data['inc_category'], data['inc_sum'],
            data['exp_category'], data['exp_sum'],
            template_path,
            file_path
        )
        await update.message.reply_document(document=open(file_path, "rb"))
        os.remove(file_path)

app = ApplicationBuilder().token(TOKEN).build()
app.add_handler(CommandHandler("start", start))
app.add_handler(CommandHandler("help", help_command))
app.add_handler(CommandHandler("doit", doit))
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
app.run_polling()