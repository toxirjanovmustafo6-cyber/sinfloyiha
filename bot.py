"""
SINF BOTI: davomat, kechikish (+SMS), sababli/sababsiz, spravka,
tug'ilgan kunlar, sinf fondi, haftalik/oylik Excel hisobot (Telegram lichkaga).
Kutubxonalar: aiogram 3, openpyxl, aiohttp
"""
import asyncio
import hashlib
import hmac
import html
import io
import os
import re
import secrets
import sqlite3
from calendar import monthrange
from datetime import date, datetime, time, timedelta
from zoneinfo import ZoneInfo

import aiohttp
from aiogram import BaseMiddleware, Bot, Dispatcher, F
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import (
    BufferedInputFile, CallbackQuery, InlineKeyboardButton as IKB,
    InlineKeyboardMarkup, KeyboardButton, Message, ReplyKeyboardMarkup,
)
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# =============================================================== SOZLAMALAR
try:
    from dotenv import load_dotenv      # .env faylidan o'qish (VS Code uchun qulay)
    load_dotenv()
except ImportError:
    pass

BOT_TOKEN = os.getenv("BOT_TOKEN", "BU_YERGA_BOT_TOKEN")      # @BotFather dan
CLASS_NAME = os.getenv("CLASS_NAME", "Sinf")
TZ = ZoneInfo("Asia/Tashkent")
DB_FILE = "davomat.db"
FUND_MONTHLY = int(os.getenv("FUND_MONTHLY", "20000"))        # oylik fond puli (so'm) - o'zgartiring
REPORT_HOUR = 18                                              # hisobot yuborish soati
BIRTHDAY_HOUR = 8                                             # tug'ilgan kun eslatmasi soati

# SMS (Eskiz.uz) - bo'sh qoldirsangiz SMS yuborilmaydi, bot qolgan hamma ishni qiladi
ESKIZ_EMAIL = os.getenv("ESKIZ_EMAIL", "")
ESKIZ_PASSWORD = os.getenv("ESKIZ_PASSWORD", "")
ESKIZ_FROM = os.getenv("ESKIZ_FROM", "4546")

# Login / parollar (bot ishga tushganda xeshlanib bazaga yoziladi)
USERS = {
    "Nuralibek": "Farangiz09",
    "Zarnigor": "Sapayeva",
    "Bekzod": "beko0409",
}

# Dars jadvali: para -> (kirish, chiqish, keyingi tanaffus daqiqada)
PARAS = {
    1: (time(8, 30), time(9, 50), 10),
    2: (time(10, 0), time(11, 20), 20),
    3: (time(11, 40), time(12, 50), 10),
    4: (time(13, 0), time(14, 0), 0),     # 14:00 da uyga ketish
}

PRESENT, ABSENT, EXCUSED = "present", "absent", "excused"
ICON = {None: "⬜", PRESENT: "✅", ABSENT: "❌", EXCUSED: "🟡"}
NEXT = {None: PRESENT, PRESENT: ABSENT, ABSENT: EXCUSED, EXCUSED: PRESENT}

# =============================================================== BAZA
db = sqlite3.connect(DB_FILE)
db.executescript("""
CREATE TABLE IF NOT EXISTS users (login TEXT PRIMARY KEY, salt TEXT, hash TEXT);
CREATE TABLE IF NOT EXISTS sessions (chat_id INTEGER PRIMARY KEY, login TEXT);
CREATE TABLE IF NOT EXISTS students (
    id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE, phone TEXT, birthday TEXT);
CREATE TABLE IF NOT EXISTS attendance (
    student_id INTEGER, day TEXT, status TEXT, PRIMARY KEY (student_id, day));
CREATE TABLE IF NOT EXISTS lates (
    id INTEGER PRIMARY KEY AUTOINCREMENT, student_id INTEGER, day TEXT, para INTEGER,
    arrived TEXT, minutes INTEGER, explained INTEGER DEFAULT 0, sms_ok INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS fund (
    id INTEGER PRIMARY KEY AUTOINCREMENT, student_id INTEGER, amount INTEGER, note TEXT, day TEXT);
CREATE TABLE IF NOT EXISTS spravkas (
    id INTEGER PRIMARY KEY AUTOINCREMENT, student_id INTEGER, from_day TEXT, to_day TEXT, note TEXT);
CREATE TABLE IF NOT EXISTS flags (key TEXT PRIMARY KEY);
""")
db.commit()


def hash_pw(pw: str, salt: str) -> str:
    return hashlib.pbkdf2_hmac("sha256", pw.encode(), bytes.fromhex(salt), 120_000).hex()


def seed_users():
    for login, pw in USERS.items():
        if not pw:
            print(f"⚠️  {login} uchun parol berilmagan — u kira olmaydi. NURALIBEK_PASS ni sozlang.")
            continue
        salt = secrets.token_hex(8)
        db.execute("INSERT OR REPLACE INTO users VALUES (?,?,?)", (login.lower(), salt, hash_pw(pw, salt)))
    db.commit()


def check_user(login: str, pw: str):
    row = db.execute("SELECT salt, hash FROM users WHERE login=?", (login.strip().lower(),)).fetchone()
    if row and hmac.compare_digest(row[1], hash_pw(pw, row[0])):
        return login.strip().lower()
    return None


def session_login(chat_id: int):
    r = db.execute("SELECT login FROM sessions WHERE chat_id=?", (chat_id,)).fetchone()
    return r[0] if r else None


def all_chats():
    return [r[0] for r in db.execute("SELECT chat_id FROM sessions")]


# =============================================================== YORDAMCHILAR
def now() -> datetime:
    return datetime.now(TZ)


def esc(s) -> str:
    return html.escape(str(s))


def fmt(n: int) -> str:
    return f"{n:,}".replace(",", " ")


def students():
    return db.execute("SELECT id, name, phone, birthday FROM students ORDER BY name").fetchall()


def student_name(sid: int) -> str:
    r = db.execute("SELECT name FROM students WHERE id=?", (sid,)).fetchone()
    return r[0] if r else "?"


def status_map(day: date) -> dict:
    return dict(db.execute("SELECT student_id, status FROM attendance WHERE day=?", (day.isoformat(),)).fetchall())


def set_status(sid: int, day: date, st: str):
    db.execute("INSERT OR REPLACE INTO attendance VALUES (?,?,?)", (sid, day.isoformat(), st))
    db.commit()


def has_spravka(sid: int, day: date) -> bool:
    return db.execute("SELECT 1 FROM spravkas WHERE student_id=? AND ? BETWEEN from_day AND to_day",
                      (sid, day.isoformat())).fetchone() is not None


def norm_phone(s: str):
    d = re.sub(r"\D", "", s or "")
    if len(d) == 9:
        d = "998" + d
    return d if len(d) == 12 and d.startswith("998") else None


def parse_dm(s: str):
    m = re.fullmatch(r"(\d{1,2})[./-](\d{1,2})(?:[./-](\d{4}))?", s.strip())
    if not m:
        return None
    d, mo, y = int(m[1]), int(m[2]), (int(m[3]) if m[3] else None)
    try:
        date(2000, mo, d)
    except ValueError:
        return None
    return d, mo, y


def parse_full_date(s: str):
    p = parse_dm(s)
    if not p:
        return None
    d, mo, y = p
    try:
        return date(y or now().year, mo, d)
    except ValueError:
        return None


def next_birthday(bd: str, today: date):
    p = parse_dm(bd)
    if not p:
        return None, None
    d, mo, y = p
    for yr in (today.year, today.year + 1):
        try:
            nb = date(yr, mo, d)
        except ValueError:
            nb = date(yr, 2, 28)
        if nb >= today:
            return nb, (nb.year - y if y else None)
    return None, None


def week_range(d: date):
    s = d - timedelta(days=d.weekday())
    return s, s + timedelta(days=5)


def month_range(d: date):
    return d.replace(day=1), d.replace(day=monthrange(d.year, d.month)[1])


def fund_balance() -> int:
    return db.execute("SELECT COALESCE(SUM(amount),0) FROM fund").fetchone()[0]


# =============================================================== SMS (Eskiz.uz)
_eskiz = {"token": None}


async def sms_send(phone: str, text: str) -> bool:
    if not (ESKIZ_EMAIL and ESKIZ_PASSWORD and phone):
        return False
    try:
        async with aiohttp.ClientSession() as s:
            for _ in range(2):
                if not _eskiz["token"]:
                    async with s.post("https://notify.eskiz.uz/api/auth/login",
                                      data={"email": ESKIZ_EMAIL, "password": ESKIZ_PASSWORD}) as r:
                        _eskiz["token"] = (await r.json()).get("data", {}).get("token")
                async with s.post("https://notify.eskiz.uz/api/message/sms/send",
                                  headers={"Authorization": f"Bearer {_eskiz['token']}"},
                                  data={"mobile_phone": phone, "message": text, "from": ESKIZ_FROM}) as r:
                    if r.status == 401:
                        _eskiz["token"] = None
                        continue
                    return r.status == 200
    except Exception as e:  # noqa
        print("SMS xato:", e)
    return False


# =============================================================== EXCEL
def build_report(start: date, end: date, title: str) -> bytes:
    days = [start + timedelta(i) for i in range((end - start).days + 1)]
    days = [d for d in days if d.weekday() != 6]           # yakshanba yo'q
    s_iso, e_iso = start.isoformat(), end.isoformat()
    hfont = Font(bold=True, color="FFFFFF")
    hfill = PatternFill("solid", fgColor="305496")
    center = Alignment(horizontal="center")
    colors = {"+": "C6EFCE", "K": "FFEB9C", "S": "BDD7EE", "-": "FFC7CE"}

    wb = Workbook()
    ws = wb.active
    ws.title = "Davomat"
    ws["A1"] = f"{CLASS_NAME} — {title} ({start:%d.%m.%Y} – {end:%d.%m.%Y})"
    ws["A1"].font = Font(bold=True, size=13)
    head = ["№", "O'quvchi"] + [f"{d:%d.%m}" for d in days] + \
           ["Keldi", "Kechikdi (marta)", "Kech (daq)", "Sababli", "Sababsiz"]
    for c, h in enumerate(head, 1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font, cell.fill, cell.alignment = hfont, hfill, center

    for i, (sid, name, _, _) in enumerate(students(), 1):
        r = 3 + i
        ws.cell(row=r, column=1, value=i)
        ws.cell(row=r, column=2, value=name)
        att = dict(db.execute("SELECT day, status FROM attendance WHERE student_id=? AND day BETWEEN ? AND ?",
                              (sid, s_iso, e_iso)).fetchall())
        late_days = {x[0] for x in db.execute(
            "SELECT DISTINCT day FROM lates WHERE student_id=? AND day BETWEEN ? AND ?", (sid, s_iso, e_iso))}
        late_min = db.execute("SELECT COALESCE(SUM(minutes),0) FROM lates WHERE student_id=? AND day BETWEEN ? AND ?",
                              (sid, s_iso, e_iso)).fetchone()[0]
        came = late_n = exc = unexc = 0
        for c, d in enumerate(days, 3):
            iso, st, sym = d.isoformat(), att.get(d.isoformat()), None
            if st == PRESENT:
                came += 1
                if iso in late_days:
                    sym, late_n = "K", late_n + 1
                else:
                    sym = "+"
            elif st == EXCUSED:
                sym, exc = "S", exc + 1
            elif st == ABSENT:
                sym, unexc = "-", unexc + 1
            cell = ws.cell(row=r, column=c)
            cell.alignment = center
            if sym:
                cell.value = sym
                cell.fill = PatternFill("solid", fgColor=colors[sym])
        base = 3 + len(days)
        for k, v in enumerate((came, late_n, late_min, exc, unexc)):
            ws.cell(row=r, column=base + k, value=v).alignment = center
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 28
    for c in range(3, len(head) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 10
    ws.cell(row=6 + len(students()), column=2,
            value="Belgilar: + keldi | K kech keldi | S sababli (spravka/ruxsat) | - sababsiz")

    # --- Kechikishlar varag'i
    w2 = wb.create_sheet("Kechikishlar")
    for c, h in enumerate(["Sana", "O'quvchi", "Para", "Keldi", "Kechikish (daq)", "Xat topshirdi", "SMS"], 1):
        cell = w2.cell(row=1, column=c, value=h)
        cell.font, cell.fill = hfont, hfill
    rows = db.execute(
        "SELECT l.day, s.name, l.para, l.arrived, l.minutes, l.explained, l.sms_ok FROM lates l "
        "JOIN students s ON s.id=l.student_id WHERE l.day BETWEEN ? AND ? ORDER BY l.day, l.id", (s_iso, e_iso))
    for r, row in enumerate(rows, 2):
        vals = [datetime.fromisoformat(row[0]).strftime("%d.%m.%Y"), row[1], f"{row[2]}-para", row[3], row[4],
                "ha" if row[5] else "yo'q", "yuborildi" if row[6] else "yuborilmadi"]
        for c, v in enumerate(vals, 1):
            w2.cell(row=r, column=c, value=v)
    for c, w in zip("ABCDEFG", (12, 28, 9, 9, 16, 14, 14)):
        w2.column_dimensions[c].width = w

    # --- Fond varag'i
    w3 = wb.create_sheet("Sinf fondi")
    for c, h in enumerate(["Sana", "Kim", "Summa (so'm)", "Izoh"], 1):
        cell = w3.cell(row=1, column=c, value=h)
        cell.font, cell.fill = hfont, hfill
    rows = db.execute(
        "SELECT f.day, COALESCE(s.name,'—'), f.amount, f.note FROM fund f LEFT JOIN students s ON s.id=f.student_id "
        "WHERE f.day BETWEEN ? AND ? ORDER BY f.day, f.id", (s_iso, e_iso)).fetchall()
    for r, row in enumerate(rows, 2):
        w3.cell(row=r, column=1, value=datetime.fromisoformat(row[0]).strftime("%d.%m.%Y"))
        w3.cell(row=r, column=2, value=row[1])
        w3.cell(row=r, column=3, value=row[2])
        w3.cell(row=r, column=4, value=row[3])
    inc = sum(x[2] for x in rows if x[2] > 0)
    exp = -sum(x[2] for x in rows if x[2] < 0)
    t = len(rows) + 3
    w3.cell(row=t, column=2, value="Kirim").font = Font(bold=True)
    w3.cell(row=t, column=3, value=inc)
    w3.cell(row=t + 1, column=2, value="Chiqim").font = Font(bold=True)
    w3.cell(row=t + 1, column=3, value=exp)
    w3.cell(row=t + 2, column=2, value="Umumiy balans").font = Font(bold=True)
    w3.cell(row=t + 2, column=3, value=fund_balance())
    for c, w in zip("ABCD", (12, 28, 16, 30)):
        w3.column_dimensions[c].width = w

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


async def send_report(bot: Bot, kind: str, chats):
    d = now().date()
    if kind == "week":
        s, e = week_range(d)
        title, fname = "Haftalik hisobot", f"haftalik_{s:%d.%m}-{e:%d.%m}.xlsx"
    else:
        s, e = month_range(d)
        title, fname = "Oylik hisobot", f"oylik_{d:%Y-%m}.xlsx"
    data = build_report(s, e, title)
    for chat in chats:
        await bot.send_document(chat, BufferedInputFile(data, filename=fname), caption=f"📊 {CLASS_NAME} — {title}")


# =============================================================== MENYU / KLAVIATURALAR
menu = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="✅ Davomat"), KeyboardButton(text="⏰ Kechikishlar")],
        [KeyboardButton(text="🎂 Tug'ilgan kunlar"), KeyboardButton(text="💰 Sinf fondi")],
        [KeyboardButton(text="📄 Spravkalar"), KeyboardButton(text="🕒 Dars jadvali")],
        [KeyboardButton(text="📊 Haftalik hisobot"), KeyboardButton(text="📅 Oylik hisobot")],
        [KeyboardButton(text="👥 O'quvchilar"), KeyboardButton(text="🚪 Chiqish")],
    ],
    resize_keyboard=True,
)


def students_kb(prefix: str) -> InlineKeyboardMarkup:
    btns = [IKB(text=n, callback_data=f"{prefix}:{sid}") for sid, n, *_ in students()]
    rows = [btns[i:i + 2] for i in range(0, len(btns), 2)]
    return InlineKeyboardMarkup(inline_keyboard=rows or [[IKB(text="(o'quvchilar yo'q)", callback_data="noop")]])


def att_kb(day: date) -> InlineKeyboardMarkup:
    st = status_map(day)
    late_ids = {r[0] for r in db.execute("SELECT DISTINCT student_id FROM lates WHERE day=?", (day.isoformat(),))}
    rows = [[IKB(text=f"{ICON[st.get(sid)]}{'⏰' if sid in late_ids else ''} {name}", callback_data=f"m:{sid}")]
            for sid, name, *_ in students()]
    rows.append([IKB(text="✅ Hammasi keldi", callback_data="all"),
                 IKB(text="⏰ Kechikish / kech keldi", callback_data="lt:menu")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def late_menu_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [IKB(text="➕ Kech keldi (kechikish yozish)", callback_data="lt:add")],
        [IKB(text="📋 Bugungi kechikishlar", callback_data="lt:today")],
        [IKB(text="📊 Hafta / oy statistikasi", callback_data="lt:stat")],
    ])


def schedule_text() -> str:
    lines = ["🕒 <b>Dars jadvali</b>"]
    for p, (a, b, br) in PARAS.items():
        tail = f"tanaffus {br} daq" if br else "keyin uyga ketish"
        lines.append(f"{p}-para: {a:%H:%M} – {b:%H:%M}  ({tail})")
    return "\n".join(lines)


async def safe_edit(c: CallbackQuery, text: str, kb=None):
    try:
        await c.message.edit_text(text, reply_markup=kb)
    except Exception:
        pass


# =============================================================== FSM
class Login(StatesGroup):
    login = State()
    password = State()


class AddSt(StatesGroup):
    waiting = State()


class LateMan(StatesGroup):
    waiting = State()


class FundIn(StatesGroup):
    waiting = State()


class FundOut(StatesGroup):
    waiting = State()


class Sprav(StatesGroup):
    waiting = State()


# =============================================================== LOGIN (middleware)
class Auth(BaseMiddleware):
    async def __call__(self, handler, event, data):
        uid = event.from_user.id
        if session_login(uid):
            return await handler(event, data)
        state: FSMContext = data["state"]
        if await state.get_state() in (Login.login.state, Login.password.state):
            return await handler(event, data)
        await state.set_state(Login.login)
        target = event.message if isinstance(event, CallbackQuery) else event
        if isinstance(event, CallbackQuery):
            await event.answer()
        await target.answer("🔐 Tizimga kirish\n\nLogin:")


dp = Dispatcher()
dp.message.outer_middleware(Auth())
dp.callback_query.outer_middleware(Auth())


@dp.message(Login.login, F.text)
async def got_login(m: Message, state: FSMContext):
    await state.update_data(login=m.text.strip())
    await state.set_state(Login.password)
    await m.answer("🔑 Parol:")


@dp.message(Login.password, F.text)
async def got_password(m: Message, state: FSMContext):
    login = (await state.get_data()).get("login", "")
    try:
        await m.delete()                       # parol chatda qolmasin
    except Exception:
        pass
    ok = check_user(login, m.text)
    if ok:
        db.execute("INSERT OR REPLACE INTO sessions VALUES (?,?)", (m.from_user.id, ok))
        db.commit()
        await state.clear()
        await m.answer(f"Xush kelibsiz, <b>{esc(login)}</b>! 👋\n{esc(CLASS_NAME)} boti.", reply_markup=menu)
    else:
        await state.set_state(Login.login)
        await m.answer("❌ Login yoki parol xato.\n\nLogin:")


@dp.message(Command("start"))
async def start(m: Message, state: FSMContext):
    await state.clear()
    await m.answer("Asosiy menyu 👇", reply_markup=menu)


@dp.message(F.text == "🚪 Chiqish")
@dp.message(Command("logout"))
async def logout(m: Message, state: FSMContext):
    await state.clear()
    db.execute("DELETE FROM sessions WHERE chat_id=?", (m.from_user.id,))
    db.commit()
    await m.answer("Tizimdan chiqdingiz. Qayta kirish uchun /start", reply_markup=None)


@dp.callback_query(F.data == "noop")
async def noop(c: CallbackQuery):
    await c.answer()


# =============================================================== DAVOMAT
@dp.message(F.text == "✅ Davomat")
async def mark(m: Message, state: FSMContext):
    await state.clear()
    if not students():
        return await m.answer("Avval o'quvchilarni qo'shing: «👥 O'quvchilar»")
    d = now().date()
    await m.answer(
        f"📆 <b>{d:%d.%m.%Y}</b>\nBosing: ⬜ → ✅ keldi → ❌ sababsiz → 🟡 sababli → ✅\n⏰ — bugun kech qolgan",
        reply_markup=att_kb(d))


@dp.callback_query(F.data.startswith("m:"))
async def toggle(c: CallbackQuery):
    sid, d = int(c.data[2:]), now().date()
    nxt = NEXT[status_map(d).get(sid)]
    if nxt == ABSENT and has_spravka(sid, d):
        nxt = EXCUSED                              # spravkasi bor — avtomatik sababli
    set_status(sid, d, nxt)
    await c.message.edit_reply_markup(reply_markup=att_kb(d))
    await c.answer()


@dp.callback_query(F.data == "all")
async def all_present(c: CallbackQuery):
    d = now().date()
    for sid, *_ in students():
        db.execute("INSERT OR IGNORE INTO attendance VALUES (?,?,?)", (sid, d.isoformat(), PRESENT))
    db.commit()
    await c.message.edit_reply_markup(reply_markup=att_kb(d))
    await c.answer("Belgilanmaganlar «keldi» qilindi")


# =============================================================== KECHIKISH
@dp.message(F.text == "⏰ Kechikishlar")
async def late_menu(m: Message, state: FSMContext):
    await state.clear()
    await m.answer("⏰ <b>Kechikishlar menyusi</b>", reply_markup=late_menu_kb())


@dp.callback_query(F.data == "lt:menu")
async def late_menu_cb(c: CallbackQuery):
    await c.message.answer("⏰ <b>Kechikishlar menyusi</b>", reply_markup=late_menu_kb())
    await c.answer()


@dp.callback_query(F.data == "lt:add")
async def late_pick_student(c: CallbackQuery):
    await safe_edit(c, "Kim kech keldi? 👇", students_kb("lt:s"))
    await c.answer()


@dp.callback_query(F.data.startswith("lt:s:"))
async def late_pick_para(c: CallbackQuery):
    sid = int(c.data.split(":")[2])
    rows = [[IKB(text=f"{p}-para ({a:%H:%M})", callback_data=f"lt:p:{sid}:{p}")] for p, (a, *_) in PARAS.items()]
    await safe_edit(c, f"<b>{esc(student_name(sid))}</b> — qaysi parag'a kech qoldi?", InlineKeyboardMarkup(inline_keyboard=rows))
    await c.answer()


@dp.callback_query(F.data.startswith("lt:p:"))
async def late_pick_time(c: CallbackQuery):
    _, _, sid, para = c.data.split(":")
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [IKB(text=f"🕐 Hozir keldi ({now():%H:%M})", callback_data=f"lt:now:{sid}:{para}")],
        [IKB(text="⌨️ Kelgan vaqtni kiritish", callback_data=f"lt:man:{sid}:{para}")],
    ])
    await safe_edit(c, f"<b>{esc(student_name(int(sid)))}</b>, {para}-para. Qachon keldi?", kb)
    await c.answer()


async def register_late(bot: Bot, sid: int, para: int, arrived: time) -> str:
    start = PARAS[para][0]
    minutes = (arrived.hour * 60 + arrived.minute) - (start.hour * 60 + start.minute)
    name = student_name(sid)
    if minutes <= 0:
        return f"✅ {esc(name)} {para}-paraga kechikmagan ({arrived:%H:%M})."
    day = now().date()
    db.execute("INSERT INTO lates (student_id, day, para, arrived, minutes) VALUES (?,?,?,?,?)",
               (sid, day.isoformat(), para, f"{arrived:%H:%M}", minutes))
    lid = db.execute("SELECT last_insert_rowid()").fetchone()[0]
    set_status(sid, day, PRESENT)
    phone = db.execute("SELECT phone FROM students WHERE id=?", (sid,)).fetchone()[0]
    sms = (f"Hurmatli {name}, siz {para}-paraga {minutes} daqiqa kechikdingiz ({arrived:%H:%M}). "
           f"Kechikishingning sababini tushuntirish xati qilib yoz.")
    ok = await sms_send(phone, sms)
    db.execute("UPDATE lates SET sms_ok=? WHERE id=?", (1 if ok else 0, lid))
    db.commit()
    if ok:
        note = f"📩 SMS yuborildi: {phone}"
    elif not phone:
        note = "⚠️ Telefon raqami kiritilmagan — SMS yuborilmadi."
    else:
        note = "⚠️ SMS yuborilmadi (Eskiz sozlanmagan yoki xato)."
    return f"⏰ <b>{esc(name)}</b> — {para}-para, <b>{minutes} daqiqa</b> kechikdi ({arrived:%H:%M}).\n{note}"


@dp.callback_query(F.data.startswith("lt:now:"))
async def late_now(c: CallbackQuery):
    _, _, sid, para = c.data.split(":")
    text = await register_late(c.bot, int(sid), int(para), now().time().replace(second=0, microsecond=0))
    await safe_edit(c, text)
    await c.answer()


@dp.callback_query(F.data.startswith("lt:man:"))
async def late_manual(c: CallbackQuery, state: FSMContext):
    _, _, sid, para = c.data.split(":")
    await state.set_state(LateMan.waiting)
    await state.update_data(sid=int(sid), para=int(para))
    await c.message.answer("Kelgan vaqtni yozing (masalan: 9.05 yoki 09:05):")
    await c.answer()


@dp.message(LateMan.waiting, F.text)
async def late_manual_time(m: Message, state: FSMContext):
    mt = re.fullmatch(r"(\d{1,2})[:.](\d{2})", m.text.strip())
    if not mt or int(mt[1]) > 23 or int(mt[2]) > 59:
        return await m.answer("Format xato. Masalan: 9.05")
    data = await state.get_data()
    await state.clear()
    text = await register_late(m.bot, data["sid"], data["para"], time(int(mt[1]), int(mt[2])))
    await m.answer(text)


def today_lates_view():
    d = now().date().isoformat()
    rows = db.execute(
        "SELECT l.id, s.name, l.para, l.arrived, l.minutes, l.explained, l.sms_ok FROM lates l "
        "JOIN students s ON s.id=l.student_id WHERE l.day=? ORDER BY l.id", (d,)).fetchall()
    if not rows:
        return "Bugun kechikkanlar yo'q 👏", None
    lines = ["📋 <b>Bugungi kechikishlar</b> (✍️ — xat topshirgan)\n"]
    kb = []
    for lid, name, para, arr, mins, ex, sms in rows:
        lines.append(f"{'✍️' if ex else '•'} {esc(name)} — {para}-para, {arr}, {mins} daq {'📩' if sms else ''}")
        kb.append([IKB(text=f"{'✅' if ex else '✍️'} {name}: xat topshirdi", callback_data=f"lt:x:{lid}")])
    return "\n".join(lines), InlineKeyboardMarkup(inline_keyboard=kb)


@dp.callback_query(F.data == "lt:today")
async def late_today(c: CallbackQuery):
    text, kb = today_lates_view()
    await safe_edit(c, text, kb)
    await c.answer()


@dp.callback_query(F.data.startswith("lt:x:"))
async def late_explained(c: CallbackQuery):
    db.execute("UPDATE lates SET explained = 1 - explained WHERE id=?", (int(c.data.split(":")[2]),))
    db.commit()
    text, kb = today_lates_view()
    await safe_edit(c, text, kb)
    await c.answer("Yangilandi")


@dp.callback_query(F.data == "lt:stat")
async def late_stat(c: CallbackQuery):
    d = now().date()
    ws, we = week_range(d)
    ms, me = month_range(d)

    def block(a, b):
        rows = db.execute(
            "SELECT s.name, COUNT(*), SUM(l.minutes) FROM lates l JOIN students s ON s.id=l.student_id "
            "WHERE l.day BETWEEN ? AND ? GROUP BY s.id ORDER BY COUNT(*) DESC, SUM(l.minutes) DESC",
            (a.isoformat(), b.isoformat())).fetchall()
        return "\n".join(f"{i}. {esc(n)} — {k} marta, {mn} daq" for i, (n, k, mn) in enumerate(rows, 1)) or "yo'q"

    await safe_edit(c, f"📊 <b>Shu hafta:</b>\n{block(ws, we)}\n\n📅 <b>Shu oy:</b>\n{block(ms, me)}",
                    late_menu_kb())
    await c.answer()


# =============================================================== TUG'ILGAN KUNLAR
def birthdays_text(limit=10) -> str:
    today = now().date()
    items = []
    for sid, name, phone, bd in students():
        if not bd:
            continue
        nb, age = next_birthday(bd, today)
        if nb:
            items.append(((nb - today).days, name, nb, age))
    if not items:
        return "Tug'ilgan kunlar kiritilmagan. «👥 O'quvchilar» orqali qo'shing."
    items.sort()
    lines = ["🎂 <b>Yaqin tug'ilgan kunlar</b>\n"]
    for days, name, nb, age in items[:limit]:
        when = "BUGUN 🎉" if days == 0 else ("ertaga" if days == 1 else f"{days} kundan keyin")
        lines.append(f"• {esc(name)} — {nb:%d.%m}" + (f" ({age} yosh)" if age else "") + f" — {when}")
    lines.append(f"\n💰 Fond balansi: {fmt(fund_balance())} so'm")
    return "\n".join(lines)


@dp.message(F.text == "🎂 Tug'ilgan kunlar")
async def birthdays(m: Message, state: FSMContext):
    await state.clear()
    await m.answer(birthdays_text())


# =============================================================== SINF FONDI
def fund_menu():
    text = (f"💰 <b>Sinf fondi</b>\nBalans: <b>{fmt(fund_balance())} so'm</b>\n"
            f"Oylik badal: {fmt(FUND_MONTHLY)} so'm")
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [IKB(text="➕ Pul qabul qilish", callback_data="fd:in"), IKB(text="➖ Xarajat", callback_data="fd:out")],
        [IKB(text="📋 Qarzdorlar (shu oy)", callback_data="fd:debt"), IKB(text="📜 Tarix", callback_data="fd:hist")],
    ])
    return text, kb


@dp.message(F.text == "💰 Sinf fondi")
async def fund(m: Message, state: FSMContext):
    await state.clear()
    text, kb = fund_menu()
    await m.answer(text, reply_markup=kb)


@dp.callback_query(F.data == "fd:in")
async def fund_in(c: CallbackQuery):
    await safe_edit(c, "Kim pul topshirdi? 👇", students_kb("fd:s"))
    await c.answer()


@dp.callback_query(F.data.startswith("fd:s:"))
async def fund_in_student(c: CallbackQuery, state: FSMContext):
    sid = int(c.data.split(":")[2])
    await state.set_state(FundIn.waiting)
    await state.update_data(sid=sid)
    await c.message.answer(f"<b>{esc(student_name(sid))}</b> — summani yozing.\n"
                           f"Masalan: <code>20000</code> yoki <code>20000 may oyi uchun</code>")
    await c.answer()


def parse_amount(text: str):
    parts = text.strip().split(maxsplit=1)
    digits = re.sub(r"\D", "", parts[0]) if parts else ""
    if not digits:
        return None, ""
    return int(digits), (parts[1] if len(parts) > 1 else "")


@dp.message(FundIn.waiting, F.text)
async def fund_in_amount(m: Message, state: FSMContext):
    amount, note = parse_amount(m.text)
    if not amount:
        return await m.answer("Summani raqam bilan yozing. Masalan: 20000")
    sid = (await state.get_data())["sid"]
    await state.clear()
    db.execute("INSERT INTO fund (student_id, amount, note, day) VALUES (?,?,?,?)",
               (sid, amount, note, now().date().isoformat()))
    db.commit()
    await m.answer(f"✅ {esc(student_name(sid))}: +{fmt(amount)} so'm qabul qilindi.\n"
                   f"💰 Balans: {fmt(fund_balance())} so'm")


@dp.callback_query(F.data == "fd:out")
async def fund_out(c: CallbackQuery, state: FSMContext):
    await state.set_state(FundOut.waiting)
    await c.message.answer("Xarajatni yozing: <code>summa izoh</code>\nMasalan: <code>150000 tug'ilgan kun torti</code>")
    await c.answer()


@dp.message(FundOut.waiting, F.text)
async def fund_out_amount(m: Message, state: FSMContext):
    amount, note = parse_amount(m.text)
    if not amount:
        return await m.answer("Summani raqam bilan yozing. Masalan: 150000 tort")
    await state.clear()
    db.execute("INSERT INTO fund (student_id, amount, note, day) VALUES (NULL,?,?,?)",
               (-amount, note or "xarajat", now().date().isoformat()))
    db.commit()
    await m.answer(f"➖ Xarajat: {fmt(amount)} so'm ({esc(note)})\n💰 Balans: {fmt(fund_balance())} so'm")


@dp.callback_query(F.data == "fd:debt")
async def fund_debt(c: CallbackQuery):
    ms, me = month_range(now().date())
    lines = []
    for sid, name, *_ in students():
        paid = db.execute("SELECT COALESCE(SUM(amount),0) FROM fund WHERE student_id=? AND amount>0 "
                          "AND day BETWEEN ? AND ?", (sid, ms.isoformat(), me.isoformat())).fetchone()[0]
        if paid < FUND_MONTHLY:
            lines.append(f"• {esc(name)} — {fmt(FUND_MONTHLY - paid)} so'm qarz")
    text = f"📋 <b>Qarzdorlar ({ms:%m.%Y})</b>\n\n" + ("\n".join(lines) if lines else "Hamma to'lagan 👏")
    await safe_edit(c, text, fund_menu()[1])
    await c.answer()


@dp.callback_query(F.data == "fd:hist")
async def fund_hist(c: CallbackQuery):
    rows = db.execute("SELECT f.day, COALESCE(s.name,'—'), f.amount, f.note FROM fund f "
                      "LEFT JOIN students s ON s.id=f.student_id ORDER BY f.id DESC LIMIT 15").fetchall()
    lines = [f"{datetime.fromisoformat(d):%d.%m} | {esc(n)} | {'+' if a > 0 else ''}{fmt(a)} | {esc(note or '')}"
             for d, n, a, note in rows]
    await safe_edit(c, "📜 <b>Oxirgi amallar</b>\n\n" + ("\n".join(lines) or "hali yo'q"), fund_menu()[1])
    await c.answer()


# =============================================================== SPRAVKALAR
@dp.message(F.text == "📄 Spravkalar")
async def sprav_menu(m: Message, state: FSMContext):
    await state.clear()
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [IKB(text="➕ Spravka qo'shish", callback_data="sp:add")],
        [IKB(text="📋 Spravkalar ro'yxati", callback_data="sp:list")],
    ])
    await m.answer("📄 <b>Spravkalar</b>\nSpravka muddatidagi qoldirilgan darslar avtomatik «sababli» bo'ladi.", reply_markup=kb)


@dp.callback_query(F.data == "sp:add")
async def sp_add(c: CallbackQuery):
    await safe_edit(c, "Kimga spravka? 👇", students_kb("sp:s"))
    await c.answer()


@dp.callback_query(F.data.startswith("sp:s:"))
async def sp_student(c: CallbackQuery, state: FSMContext):
    sid = int(c.data.split(":")[2])
    await state.set_state(Sprav.waiting)
    await state.update_data(sid=sid)
    await c.message.answer(
        f"<b>{esc(student_name(sid))}</b> — muddatni yozing:\n"
        f"<code>05.11 07.11 shifokor spravkasi</code>\n(bir kunlik bo'lsa faqat bitta sana)")
    await c.answer()


@dp.message(Sprav.waiting, F.text)
async def sp_save(m: Message, state: FSMContext):
    tok = m.text.split()
    d1 = parse_full_date(tok[0]) if tok else None
    if not d1:
        return await m.answer("Sana xato. Masalan: 05.11 07.11 shifokor spravkasi")
    d2 = parse_full_date(tok[1]) if len(tok) > 1 else None
    note = " ".join(tok[2:] if d2 else tok[1:])
    d2 = d2 or d1
    if d2 < d1:
        d1, d2 = d2, d1
    sid = (await state.get_data())["sid"]
    await state.clear()
    db.execute("INSERT INTO spravkas (student_id, from_day, to_day, note) VALUES (?,?,?,?)",
               (sid, d1.isoformat(), d2.isoformat(), note))
    n = db.execute("UPDATE attendance SET status=? WHERE student_id=? AND status=? AND day BETWEEN ? AND ?",
                   (EXCUSED, sid, ABSENT, d1.isoformat(), d2.isoformat())).rowcount
    db.commit()
    await m.answer(f"✅ Spravka saqlandi: {esc(student_name(sid))}, {d1:%d.%m}–{d2:%d.%m}.\n"
                   f"Sababli deb o'zgartirilgan kunlar: {n}")


@dp.callback_query(F.data == "sp:list")
async def sp_list(c: CallbackQuery):
    rows = db.execute("SELECT s.name, p.from_day, p.to_day, p.note FROM spravkas p JOIN students s ON s.id=p.student_id "
                      "ORDER BY p.id DESC LIMIT 15").fetchall()
    lines = [f"• {esc(n)}: {datetime.fromisoformat(a):%d.%m}–{datetime.fromisoformat(b):%d.%m} {esc(note or '')}"
             for n, a, b, note in rows]
    await safe_edit(c, "📋 <b>Spravkalar</b>\n\n" + ("\n".join(lines) or "hali yo'q"))
    await c.answer()


# =============================================================== DARS JADVALI
@dp.message(F.text == "🕒 Dars jadvali")
async def schedule(m: Message, state: FSMContext):
    await state.clear()
    await m.answer(schedule_text())


# =============================================================== HISOBOTLAR
@dp.message(F.text == "📊 Haftalik hisobot")
async def weekly(m: Message, state: FSMContext):
    await state.clear()
    await send_report(m.bot, "week", [m.from_user.id])


@dp.message(F.text == "📅 Oylik hisobot")
async def monthly(m: Message, state: FSMContext):
    await state.clear()
    await send_report(m.bot, "month", [m.from_user.id])


# =============================================================== O'QUVCHILAR
@dp.message(F.text == "👥 O'quvchilar")
async def students_menu(m: Message, state: FSMContext):
    await state.clear()
    rows = students()
    lines = [f"{i}. {esc(n)}" + (f" — {p}" if p else " — ❗telefon yo'q") + (f" — 🎂 {b}" if b else "")
             for i, (_, n, p, b) in enumerate(rows, 1)]
    kb = [[IKB(text="➕ Qo'shish / yangilash", callback_data="st:add")]]
    kb += [[IKB(text=f"🗑 {n}", callback_data=f"st:del:{sid}")] for sid, n, *_ in rows]
    await m.answer(f"👥 <b>O'quvchilar: {len(rows)}</b>\n\n" + ("\n".join(lines) or "hali yo'q"),
                   reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))


@dp.callback_query(F.data == "st:add")
async def st_add(c: CallbackQuery, state: FSMContext):
    await state.set_state(AddSt.waiting)
    await c.message.answer(
        "Har bir o'quvchini yangi qatordan yuboring:\n"
        "<code>Ism Familiya; telefon; tug'ilgan kun</code>\n\n"
        "Masalan:\n<code>Zarnigor Sapayeva; 901234567; 15.03.2010\nBekzod Aliyev; +998931112233; 02.09</code>\n\n"
        "Telefon va tug'ilgan kun ixtiyoriy. Mavjud o'quvchini qayta yuborsangiz, ma'lumoti yangilanadi.\n"
        "Telefon — kechikishda SMS borishi uchun kerak.")
    await c.answer()


@dp.message(AddSt.waiting, F.text)
async def st_save(m: Message, state: FSMContext):
    n = 0
    for line in m.text.splitlines():
        parts = [p.strip() for p in re.split(r"\s*;\s*", line) if p.strip()]
        if not parts:
            continue
        phone = bd = None
        for p in parts[1:]:
            if norm_phone(p):
                phone = norm_phone(p)
            elif parse_dm(p):
                d, mo, y = parse_dm(p)
                bd = f"{d:02d}.{mo:02d}" + (f".{y}" if y else "")
        db.execute("""INSERT INTO students (name, phone, birthday) VALUES (?,?,?)
                      ON CONFLICT(name) DO UPDATE SET phone=COALESCE(excluded.phone, phone),
                      birthday=COALESCE(excluded.birthday, birthday)""", (parts[0], phone, bd))
        n += 1
    db.commit()
    await state.clear()
    await m.answer(f"✅ {n} ta o'quvchi saqlandi.", reply_markup=menu)


@dp.callback_query(F.data.startswith("st:del:"))
async def st_del(c: CallbackQuery):
    sid = int(c.data.split(":")[2])
    for t in ("attendance", "lates", "spravkas"):
        db.execute(f"DELETE FROM {t} WHERE student_id=?", (sid,))
    db.execute("DELETE FROM students WHERE id=?", (sid,))
    db.commit()
    await c.answer("O'chirildi")
    await c.message.delete()


# =============================================================== AVTOMATIK ISHLAR
def once(key: str) -> bool:
    cur = db.execute("INSERT OR IGNORE INTO flags VALUES (?)", (key,))
    db.commit()
    return cur.rowcount == 1


async def scheduler(bot: Bot):
    while True:
        try:
            t = now()
            today = t.date()
            # tug'ilgan kun eslatmasi (bugun / ertaga / 3 kun qoldi)
            if t.hour == BIRTHDAY_HOUR and once(f"bd-{today}"):
                lines = []
                for _, name, _, bd in students():
                    if not bd:
                        continue
                    nb, age = next_birthday(bd, today)
                    left = (nb - today).days if nb else None
                    if left in (0, 1, 3):
                        when = {0: "BUGUN", 1: "ertaga", 3: "3 kundan keyin"}[left]
                        lines.append(f"🎂 {esc(name)} — {when} tug'ilgan kun!" + (f" ({age} yosh)" if age else ""))
                if lines:
                    text = "\n".join(lines) + f"\n\n💰 Fond balansi: {fmt(fund_balance())} so'm"
                    for chat in all_chats():
                        await bot.send_message(chat, text)
            # haftalik (shanba) va oylik (oy oxiri) hisobot
            if t.hour == REPORT_HOUR:
                if today.weekday() == 5 and once(f"w-{today}"):
                    await send_report(bot, "week", all_chats())
                if (today + timedelta(days=1)).day == 1 and once(f"m-{today}"):
                    await send_report(bot, "month", all_chats())
        except Exception as e:  # noqa
            print("scheduler xato:", e)
        await asyncio.sleep(60)


async def main():
    seed_users()
    bot = Bot(BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    asyncio.create_task(scheduler(bot))
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
