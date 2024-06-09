from telegram import Update, InputFile
from telegram.ext import Application, CommandHandler, ContextTypes
import requests
import pandas as pd
from io import BytesIO
import asyncio
import nest_asyncio

# Замените YOUR_TELEGRAM_BOT_TOKEN на токен вашего Телеграм-бота
TOKEN = '7044724125:AAGRUbChusByOgnkh2q_sW4nCkr5tlEP-dU'
# Замените URL на URL вашего веб-приложения Google Apps Script
GOOGLE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwsgfpXgO8hleKGpdxD3qCNGOOSVO4JM89v6SgJr1KlGkaZuQ9ae3mnlHtfG3BWGFZR/exec'

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text('Привет! Используйте /report для получения отчета.\nИспользуйте /help для получения списка доступных команд.')

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    help_text = (
        "/start - Запустить бота\n"
        "/report - Получить последний отчет\n"
        "/report_day ГГГГ-ММ-ДД - Получить отчет за конкретный день\n"
        "/report_period ГГГГ-ММ-ДД ГГГГ-ММ-ДД - Получить отчет за конкретный период\n"
        "/stats ГГГГ-ММ-ДД - Получить статистику заказов (год, месяц, день)\n"
        "/export_excel - Экспортировать отчет в формате Excel"
    )
    await update.message.reply_text(help_text)

async def fetch_report(url: str) -> str:
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        report_text = ''
        for entry in data:
            report_text += (
                f"Имя: {entry['Name']}\n"
                f"Email: {entry['Email']}\n"
                f"Телефон: {entry['Phone']}\n"
                f"ID заказа: {entry['orderid']}\n"
                f"Продукты: {entry['products']}\n"
                f"Цена: {entry['price']} {entry['Currency']}\n"
                f"Отправлено: {entry['sent']}\n\n"
            )
        return report_text
    except requests.exceptions.RequestException as e:
        return f"Ошибка HTTP-запроса: {e}"
    except ValueError:
        return "Не удалось декодировать ответ JSON."

async def report(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    report_text = await fetch_report(GOOGLE_SCRIPT_URL)
    await update.message.reply_text(report_text)

async def report_day(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if context.args:
        date = context.args[0]
        report_text = await fetch_report(f"{GOOGLE_SCRIPT_URL}?date={date}")
        await update.message.reply_text(report_text)
    else:
        await update.message.reply_text('Пожалуйста, укажите дату в формате ГГГГ-ММ-ДД.')

async def report_period(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if len(context.args) == 2:
        start_date, end_date = context.args
        report_text = await fetch_report(f"{GOOGLE_SCRIPT_URL}?start_date={start_date}&end_date={end_date}")
        await update.message.reply_text(report_text)
    else:
        await update.message.reply_text('Пожалуйста, укажите начальную и конечную даты в формате ГГГГ-ММ-ДД.')

async def stats(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if context.args:
        date = context.args[0]
        response = requests.get(GOOGLE_SCRIPT_URL)
        if response.status_code == 200:
            data = response.json()
            year_count = sum(1 for entry in data if entry['sent'].startswith(date[:4]))
            month_count = sum(1 for entry in data if entry['sent'].startswith(date[:7]))
            day_count = sum(1 for entry in data if entry['sent'].startswith(date))
            stats_text = (
                f"Заказы за {date[:4]}: {year_count}\n"
                f"Заказы за {date[:7]}: {month_count}\n"
                f"Заказы за {date}: {day_count}\n"
            )
            await update.message.reply_text(stats_text)
        else:
            await update.message.reply_text('Не удалось получить данные из Google Sheets.')
    else:
        await update.message.reply_text('Пожалуйста, укажите дату в формате ГГГГ-ММ-ДД.')

async def export_excel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    response = requests.get(GOOGLE_SCRIPT_URL)
    if response.status_code == 200:
        data = response.json()
        df = pd.DataFrame(data)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Заказы')
        output.seek(0)
        await update.message.reply_document(InputFile(output, filename='orders.xlsx'))
    else:
        await update.message.reply_text('Не удалось получить данные из Google Sheets.')

async def main() -> None:
    application = Application.builder().token(TOKEN).build()

    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler("report", report))
    application.add_handler(CommandHandler("report_day", report_day))
    application.add_handler(CommandHandler("report_period", report_period))
    application.add_handler(CommandHandler("stats", stats))
    application.add_handler(CommandHandler("export_excel", export_excel))

    await application.run_polling()

if __name__ == '__main__':
    nest_asyncio.apply()
    asyncio.run(main())
