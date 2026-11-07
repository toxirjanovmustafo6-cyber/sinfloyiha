# Sinf loyihasi (VS Code uchun)

- `bot.py` — Telegram bot (davomat, kechikish + SMS, fond, spravka, tug'ilgan kun, Excel hisobot)
- `web/sinf_davomat.html` — vizual sahifa

## VS Code'da ishga tushirish
1. VS Code → File → Open Folder → shu papkani oching. Taklif qilingan kengaytmalarni (Python, Live Server) o'rnating.
2. Terminal (Ctrl+`) ochib:
   ```
   python -m venv .venv
   .venv\Scripts\activate        (Mac/Linux: source .venv/bin/activate)
   pip install -r requirements.txt
   ```
3. `.env.example` ni `.env` deb nusxalang va BOT_TOKEN ni yozing (@BotFather dan).
4. F5 bosing ("Botni ishga tushirish") yoki terminalda `python bot.py`.
5. Telegramda botga /start yozing. Login: Nuralibek / Zarnigor / Bekzod.

## Vizual sahifa
`web/sinf_davomat.html` ustida sichqonchaning o'ng tugmasi → "Open with Live Server".
(Bu holda ma'lumot faqat shu brauzerda saqlanadi.)
