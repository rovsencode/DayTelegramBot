import asyncio
import telegram
import openpyxl
from datetime import datetime
import pytz

# 🔹 Azərbaycan saat qurşağını təyin edirik
baku_tz = pytz.timezone("Asia/Baku")

# 🔹 Telegram Bot Məlumatları
TOKEN = "7794374904:AAE8SLnOpyl1unO-1w8OXrLJETB4aAN8b0E"
CHAT_ID = "7078657063"

bot = telegram.Bot(token=TOKEN)

# 🔹 Excel faylından bugünkü tapşırığı oxuyan funksiya
def get_today_task():
    file_path = "6_Həftəlik_IT_Help_Desk_İnkişaf_Planı.xlsx"
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    today = datetime.now(baku_tz).strftime("%d-%m-%Y")

    for row in ws.iter_rows(min_row=2, values_only=True):
        week, date, topic, task, _, _ = row
        if date == today:
            return f"📅 {date}\n📌 Mövzu: {topic}\n✅ Tapşırıq və resurslar: {task}"
    
    return "📌 Bu gün üçün tapşırıq yoxdur."

# 🔹 Hər gün saat 10:00-da mesaj göndərən asinxron funksiya
async def send_daily_task():
    while True:
        now = datetime.now(baku_tz)
        if now.hour == 14 and now.minute==0:  # 10:00-da göndər
            message = get_today_task()
            await bot.send_message(chat_id=CHAT_ID, text=message)
        
        await asyncio.sleep(60)  # Hər 60 saniyədən bir yoxlayır

# 🔹 Botu işə salan əsas funksiya
async def main():
    await send_daily_task()

if __name__ == "__main__":
    asyncio.run(main())  # Asinxron funksiyanı işə salır
