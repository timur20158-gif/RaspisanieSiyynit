from datetime import datetime, timedelta
import hashlib
import html
import json
import os
import queue
import re
import shutil
import subprocess
from contextlib import contextmanager
import sqlite3
import tempfile
import threading
import time
from zoneinfo import ZoneInfo

import docx
import telebot
from telebot.apihelper import ApiTelegramException
from telebot.types import (
    BotCommand,
    BotCommandScopeAllGroupChats,
    BotCommandScopeChat,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    InlineQueryResultArticle,
    InlineQueryResultCachedDocument,
    InlineQueryResultCachedPhoto,
    InputMediaPhoto,
    InputTextMessageContent,
    KeyboardButton,
    ReplyKeyboardMarkup,
)
from apscheduler.schedulers.background import BackgroundScheduler

# --- НАСТРОЙКИ ---
# Укажи актуальный токен в переменной окружения BOT_TOKEN или вставь вместо 'YOUR_BOT_TOKEN_HERE'
TOKEN = os.getenv('BOT_TOKEN', 'YOUR_BOT_TOKEN_HERE')
ADMIN_USERNAME = 'tyamuuuuuuuuur'
LOCAL_TZ = ZoneInfo('Asia/Yekaterinburg')
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.getenv('DATA_DIR') or ('/data' if os.path.isdir('/data') else BASE_DIR)
DB_PATH = os.path.join(DATA_DIR, 'schedule.db')
SCHEDULE_JSON_PATH = os.path.join(DATA_DIR, 'schedule.json')
SEMESTER_META_PATH = os.path.join(DATA_DIR, 'semester_meta.json')

bot = telebot.TeleBot(TOKEN, threaded=True, num_threads=8)

DAYS_RUS = [
    'Понедельник',
    'Вторник',
    'Среда',
    'Четверг',
    'Пятница',
    'Суббота',
    'Воскресенье',
]

# Жёсткая привязка времени к реальному номеру пары
PAIR_NUMBERS = {
    '08:00 - 09:20': 1,
    '09:35 - 10:55': 2,
    '11:10 - 12:30': 3,
    '13:30 - 14:50': 4,
    '15:05 - 16:25': 5,
    '16:40 - 18:00': 6,
}

NUM_TO_TIME_SLOT = {v: k for k, v in PAIR_NUMBERS.items()}

PAIR_EMOJIS = {
    1: '1️⃣',
    2: '2️⃣',
    3: '3️⃣',
    4: '4️⃣',
    5: '5️⃣',
    6: '6️⃣',
}

# --- СПРАВОЧНИК ФАКУЛЬТЕТОВ СИ УУНиТ ---
FACULTIES_META = {
    'ЕМФ': {
        'short': 'ЕМФ',
        'full': 'Естественно-математический факультет',
        'genitive': 'Естественно-математического факультета',
        'emoji': '🧬',
    },
    'ПФ': {
        'short': 'Педфак',
        'full': 'Педагогический факультет',
        'genitive': 'Педагогического факультета',
        'emoji': '📚',
    },
    'ФЭП': {
        'short': 'ФЭП',
        'full': 'Факультет экономики и права',
        'genitive': 'Факультета экономики и права',
        'emoji': '⚖️',
    },
    'ТФ': {
        'short': 'ТФ',
        'full': 'Технологический факультет',
        'genitive': 'Технологического факультета',
        'emoji': '⚙️',
    },
}

FACULTY_ORDER = ['ЕМФ', 'ПФ', 'ФЭП', 'ТФ']

# Базовые практики (дополняются автоматически из загружаемых .docx файлов)
DEFAULT_PRACTICE_WEEKS = {}

# Динамические структуры (полностью формируются из schedule.json и semester_meta.json)
FACULTY_STRUCTURE = {fac: {} for fac in FACULTY_ORDER}
PRACTICE_WEEKS = {k: set(v) for k, v in DEFAULT_PRACTICE_WEEKS.items()}

# Старт 1-й учебной недели учебного года (понедельник недели 1 сентября, сквозной счёт без сброса во 2-м семестре)
ACADEMIC_YEAR_START_DATE = datetime(2026, 8, 31, tzinfo=LOCAL_TZ)

PENDING_DOCX_DIR = os.path.join(tempfile.gettempdir(), 'si_uunit_docx_imports')
os.makedirs(PENDING_DOCX_DIR, exist_ok=True)
pending_docx_imports = {}


def atomic_write_json(filepath, data):
    dir_name = os.path.dirname(os.path.abspath(filepath)) or '.'
    os.makedirs(dir_name, exist_ok=True)
    fd, tmp_path = tempfile.mkstemp(dir=dir_name, prefix='.tmp_', suffix='.json')
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp_path, filepath)
    except Exception:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass
        raise


def save_pending_docx_import(token, data_dict):
    pending_docx_imports[token] = data_dict
    serializable = dict(data_dict)
    if isinstance(serializable.get('year_start_dt'), datetime):
        serializable['year_start_dt'] = serializable['year_start_dt'].strftime('%d.%m.%Y')
    try:
        atomic_write_json(os.path.join(PENDING_DOCX_DIR, f'{token}.json'), serializable)
    except Exception as e:
        print(f'Ошибка сохранения сессии импорта: {e}')


def get_pending_docx_import(token, pop=False):
    item = pending_docx_imports.pop(token, None) if pop else pending_docx_imports.get(token)
    fpath = os.path.join(PENDING_DOCX_DIR, f'{token}.json')
    if item is None and os.path.exists(fpath):
        try:
            with open(fpath, 'r', encoding='utf-8') as f:
                item = json.load(f)
            if isinstance(item.get('year_start_dt'), str):
                item['year_start_dt'] = datetime.strptime(item['year_start_dt'], '%d.%m.%Y').replace(tzinfo=LOCAL_TZ)
            if not pop:
                pending_docx_imports[token] = item
        except Exception as e:
            print(f'Ошибка чтения сессии импорта: {e}')
            item = None
    if pop and os.path.exists(fpath):
        try:
            os.remove(fpath)
        except Exception:
            pass
    return item

# Буфер для сборки альбомов (media_group) при добавлении ДЗ
media_group_buffer = {}
media_group_lock = threading.Lock()

# --- ФОНОВАЯ ОЧЕРЕДЬ РАССЫЛКИ УВЕДОМЛЕНИЙ ---
MAILING_QUEUE = queue.Queue()
MAILING_INTERVAL_SEC = 0.06


def _requeue_after_delay(task, delay_sec):
    def _put_back():
        MAILING_QUEUE.put(task)
    t = threading.Timer(max(1.0, float(delay_sec)), _put_back)
    t.daemon = True
    t.start()


def _mailing_worker():
    while True:
        task = MAILING_QUEUE.get()
        if task is None:
            MAILING_QUEUE.task_done()
            break
        try:
            kind = task.get('kind', 'text')
            user_id = task['user_id']
            retries = int(task.get('_retries', 0))
            if kind == 'media':
                status, retry_after = _direct_send_hw_media(
                    user_id,
                    task.get('photo_ids') or [],
                    task.get('caption_html', ''),
                    reply_markup=task.get('reply_markup'),
                    doc_file_id=task.get('doc_file_id', ''),
                    doc_file_name=task.get('doc_file_name', ''),
                )
            else:
                status, retry_after = _direct_send_message(
                    user_id,
                    task.get('text', ''),
                    reply_markup=task.get('reply_markup'),
                    parse_mode=task.get('parse_mode', 'HTML'),
                )
            if status == 'retry' and retries < 3:
                requeued_task = dict(task)
                requeued_task['_retries'] = retries + 1
                _requeue_after_delay(requeued_task, retry_after)
        except Exception as e:
            print(f'Ошибка в фоновом потоке рассылки: {e}')
        finally:
            MAILING_QUEUE.task_done()
            time.sleep(MAILING_INTERVAL_SEC)


_mailing_thread = threading.Thread(target=_mailing_worker, name='MailingWorker', daemon=True)
_mailing_thread.start()


def enqueue_message(user_id, text, reply_markup=None, parse_mode='HTML'):
    MAILING_QUEUE.put({
        'kind': 'text',
        'user_id': user_id,
        'text': text,
        'reply_markup': reply_markup,
        'parse_mode': parse_mode,
        '_retries': 0,
    })


def enqueue_hw_media(user_id, photo_ids, caption_html, reply_markup=None, doc_file_id='', doc_file_name=''):
    MAILING_QUEUE.put({
        'kind': 'media',
        'user_id': user_id,
        'photo_ids': photo_ids,
        'caption_html': caption_html,
        'reply_markup': reply_markup,
        'doc_file_id': doc_file_id,
        'doc_file_name': doc_file_name,
        '_retries': 0,
    })


# --- ЕДИНАЯ ФУНКЦИЯ ЛОКАЛЬНОГО ВРЕМЕНИ (UTC+5) ---
def now_local():
    return datetime.now(LOCAL_TZ)


def get_academic_year_start_by_year(start_year):
    sept_1 = datetime(int(start_year), 9, 1, tzinfo=LOCAL_TZ)
    # Понедельник той недели, в которой находится 1 сентября (для 2026 это 31.08.2026, для 2027 это 30.08.2027)
    return sept_1 - timedelta(days=sept_1.weekday())


def extract_academic_start_from_title(title_text):
    txt = str(title_text or '')
    m = re.search(r'(20\d{2})\s*[/\\–—-]\s*(20\d{2})', txt)
    if m:
        y1, y2 = int(m.group(1)), int(m.group(2))
        if y2 == y1 + 1:
            return get_academic_year_start_by_year(y1)
    return None


def get_default_academic_year_start(ref_dt=None):
    if ref_dt is None:
        ref_dt = now_local()
    # В июле уже публикуют расписание на новый учебный год (сентябрь), поэтому границей считаем июнь (<= 6).
    start_year = ref_dt.year - 1 if ref_dt.month <= 6 else ref_dt.year
    return get_academic_year_start_by_year(start_year)



def get_faculty_short(fac_code):
    meta = FACULTIES_META.get(fac_code)
    return meta['short'] if meta else str(fac_code or 'ЕМФ')


def get_faculty_full(fac_code):
    meta = FACULTIES_META.get(fac_code)
    return meta['full'] if meta else str(fac_code or 'Естественно-математический факультет')


def detect_faculty_from_text(text):
    low = str(text or '').lower()
    if any(k in low for k in ('педагогического факультета', 'педагогический факультет', 'педфак', 'офо_пф', 'офо пф', '_пф_', ' пф ')):
        return 'ПФ'
    if any(k in low for k in ('экономики и права', 'факультет экономики', 'фэп', 'офо_фэп', '_фэп_')):
        return 'ФЭП'
    if any(k in low for k in ('технологического факультета', 'технологический факультет', 'техфак', 'офо_тф', '_тф_')):
        return 'ТФ'
    if any(k in low for k in ('естественно-математического', 'естественно-математический', 'емф', 'офо_емф', '_емф_')):
        return 'ЕМФ'
    return None


def infer_faculty_by_group(group_name):
    for fac_code, courses in FACULTY_STRUCTURE.items():
        for crs, groups in courses.items():
            if group_name in groups:
                return fac_code
    low = str(group_name or '').lower()
    if any(k in low for k in ('экономик', 'менеджмент', 'юриспруденц', 'право', 'финанс', 'бухгалтер', 'государственное и муниципальное')):
        return 'ФЭП'
    if any(k in low for k in ('технолог', 'агроинженер', 'пожарн', 'безопасность техносфер', 'строительств', 'машиностроен', 'электроэнерг')):
        return 'ТФ'
    if any(k in low for k in ('иностранный язык', 'психолог', 'дошкольн', 'начальн', 'логопед', 'дефектолог')):
        return 'ПФ'
    return 'ЕМФ'


def is_group_valid(faculty_code, course, group_name):
    courses = FACULTY_STRUCTURE.get(faculty_code, {})
    return group_name in courses.get(course, [])


# --- УНИВЕРСАЛЬНЫЙ ПАРСЕР УЧЕБНОГО РАСПИСАНИЯ ИЗ .DOCX / .DOC (ФОРМАТ СИ УУНиТ) ---
W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# Блок недель и типа занятия: (2-6, 8-15 – пз), (1-9 ЛК поток ПФ+ТФ+ФЭП), (1-8 1 п/г ЛК, 1-8 2 п/г ПЗ)
WEEK_BLOCK_RE = re.compile(
    r'\(\s*(\d+(?:\s*[-–—]\s*\d+)?(?:\s*,\s*\d+(?:\s*[-–—]\s*\d+)?)*)\s*'
    r'([-–—]?\s*(?:(?:\d+|[IІVХX]+)\s*(?:п/г|подгр(?:уппа)?\.?)|ЛК|ПЗ|Лаб|Лвб|ЛР|ПР|Сем|поток)[^)]*)\)',
    re.IGNORECASE,
)

# Аудитория: "ауд. 206", "ауд. 206а", "ауд. 301-б", "ауд. 421, 423", "ауд. конф-зал", "215 ауд.", "417а ауд."
ROOM_RE = re.compile(
    r'(?:ауд\.?\s*(?:конф[-\s]*[зэ]ал|спорт[-\s]*зал|акт(?:овый)?[-\s]*зал|[№#]?\s*\d+(?:\s*[-/]?[а-яёa-z])?(?:\s*,\s*\d+(?:\s*[-/]?[а-яёa-z])?)*)|\b\d{1,4}(?:\s*[-/]?[а-яёa-z])?\s*ауд\.?)',
    re.IGNORECASE,
)

# Полное упоминание преподавателя: [Степень/должность] Фамилия И.О.
TEACHER_FULL_RE = re.compile(
    r'(?:'
    r'(?:(?:канд\.|к\.|д-р\.?|д\.)\s*(?:[а-яё\-]+\.?\s*)*н\.?,?\s*|доц\.?,?\s*|проф\.?,?\s*|ст\.?\s*преп\.?,?\s*|ст\.?\s*пр\.?,?\s*|преп\.?,?\s*|асс\.?,?\s*)+'
    r')?'
    r'[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.',
    re.IGNORECASE,
)

# Поддержка времени как с разделителем ('08:00 - 09:20'), так и слитно ('0800 – 0920')
TIME_NORM_RE = re.compile(
    r'(\d{1,2})\s*[:.\-]?\s*(\d{2})\s*[-–—]\s*(\d{1,2})\s*[:.\-]?\s*(\d{2})'
)

DAY_MAP = {
    'понедельник': 'Понедельник',
    'вторник': 'Вторник',
    'среда': 'Среда',
    'четверг': 'Четверг',
    'пятница': 'Пятница',
    'суббота': 'Суббота',
    'воскресенье': 'Воскресенье',
}


def normalize_room_str(raw_room):
    s = str(raw_room or '').strip()
    if not s:
        return ''
    s = s.replace('конф-эал', 'конф-зал')
    m_rev = re.match(r'^(\d{1,4}(?:\s*[-/]?[а-яёa-z])?)\s*ауд\.?$', s, re.IGNORECASE)
    if m_rev:
        return f'ауд. {m_rev.group(1)}'
    s = re.sub(r'^ауд\.?\s*', 'ауд. ', s, flags=re.IGNORECASE)
    s = re.sub(r'\s*,\s*', ', ', s).strip()
    return s


def clean_teacher_str(raw_teacher):
    s = re.sub(r'\s+', ' ', str(raw_teacher or '')).strip(' ,;-/\\.')
    s = re.sub(r'^\d+\s*', '', s).strip(' ,;-/\\.')
    if not s or len(s) < 3 or s.isdigit():
        return ''
    return s


def get_tc_clean_text(tc):
    if tc is None:
        return ''
    paras = []
    for p in tc.findall(f'{W_NS}p'):
        runs_txt = [t.text for t in p.iter(f'{W_NS}t') if t.text]
        line = ''.join(runs_txt).strip()
        if line:
            paras.append(line)
    return '\n'.join(paras)


def build_logical_grid(table):
    """Строит корректную 2D-сетку ячеек таблицы Word с учётом gridSpan и vMerge."""
    grid = []
    for tr in table.rows:
        row_tcs = []
        col_cursor = 0
        trPr = tr._tr.find(f'{W_NS}trPr')
        if trPr is not None:
            gb = trPr.find(f'{W_NS}gridBefore')
            if gb is not None:
                gb_val = int(gb.get(f'{W_NS}val', 0))
                for _ in range(gb_val):
                    row_tcs.append(None)
                    col_cursor += 1

        for tc in tr._tr.findall(f'{W_NS}tc'):
            tcPr = tc.find(f'{W_NS}tcPr')
            grid_span = 1
            v_merge = None
            if tcPr is not None:
                gs = tcPr.find(f'{W_NS}gridSpan')
                if gs is not None:
                    grid_span = int(gs.get(f'{W_NS}val', 1))
                vm = tcPr.find(f'{W_NS}vMerge')
                if vm is not None:
                    v_merge = vm.get(f'{W_NS}val', 'continue')

            for _ in range(grid_span):
                if v_merge == 'continue' and grid and col_cursor < len(grid[-1]):
                    row_tcs.append(grid[-1][col_cursor])
                else:
                    row_tcs.append(tc)
                col_cursor += 1
        grid.append(row_tcs)
    return grid


def normalize_time_slot(raw_time):
    m = TIME_NORM_RE.search(str(raw_time or ''))
    if not m:
        return None
    h1, m1, h2, m2 = m.groups()
    h1_i, m1_i, h2_i, m2_i = int(h1), int(m1), int(h2), int(m2)
    if not (0 <= h1_i <= 23 and 0 <= h2_i <= 23 and 0 <= m1_i <= 59 and 0 <= m2_i <= 59):
        return None
    return f'{h1_i:02d}:{m1} - {h2_i:02d}:{m2}'


def normalize_day_name(raw_day):
    cleaned = re.sub(r'[^а-яА-ЯёЁ]', '', str(raw_day or '')).lower()
    return DAY_MAP.get(cleaned)


def parse_column_header(header_text):
    """
    Универсально распознаёт курс, название направления/профиля и недели практики
    из заголовка столбца таблицы Word для любого факультета.
    """
    flat = re.sub(r'\s+', ' ', str(header_text or '')).strip()
    if not flat:
        return None, None, None

    crs_m = re.search(r'(\d)\s*курс', flat, re.IGNORECASE)
    if not crs_m:
        return None, None, None
    course = int(crs_m.group(1))

    prac_m = re.search(r'практика\s*\((\d+)\s*[-–—]\s*(\d+)\)', flat, re.IGNORECASE)
    if not prac_m:
        prac_m = re.search(r'\((\d+)\s*[-–—]\s*(\d+)\s*нед\.?\s*[-–—]\s*практика\)', flat, re.IGNORECASE)
    practice_range = (int(prac_m.group(1)), int(prac_m.group(2))) if prac_m else None

    name = re.sub(r'\d\s*курс', '', flat, flags=re.IGNORECASE)
    name = re.sub(r'практика\s*\([^)]*\)', '', name, flags=re.IGNORECASE)
    name = re.sub(r'\(\d+\s*[-–—]\s*\d+\s*нед\.?\s*[-–—]\s*практика\)', '', name, flags=re.IGNORECASE)
    name = re.sub(r'Направление\s*(?:подготовки)?\s*', '', name, flags=re.IGNORECASE)
    name = re.sub(r'\d{2}\.\d{2}\.\d{2}\s*', '', name)
    name = re.sub(r'\(с\s+двумя\s+профилями(?:\s+подготовки)?\)', '', name, flags=re.IGNORECASE)
    name = re.sub(r'\b\d+\s*[-–—]\s*\d+\s*нед\.?\s*', '', name, flags=re.IGNORECASE)
    name = re.sub(r'(?:ауд\.?\s*\d+|\d+\s*ауд\.?)', '', name, flags=re.IGNORECASE)
    name = re.sub(r'\s+', ' ', name).strip(' ,;')

    prof_m = re.search(r'\((?:направленность\s*)?профиль\s+([^)]+)\)', name, re.IGNORECASE)
    if not prof_m:
        prof_m = re.search(r'направленность\s*\(профиль\)\s*(.+)$', name, re.IGNORECASE)
    profile_raw = prof_m.group(1).strip(' «»"') if prof_m else ''
    base_dir = re.sub(r'\(?(?:направленность\s*)?\(?профиль\)?.*$', '', name, flags=re.IGNORECASE).strip(' ,;')

    low_base = base_dir.lower()
    low_prof = profile_raw.lower()

    if 'прикладная математика' in low_base:
        group_title = f'Прикладная математика и информатика ({course} курс)'
    elif 'экология' in low_base:
        group_title = f'Экология и природопользование ({course} курс)'
    elif 'биология' in low_base and 'педагогическое' not in low_base:
        group_title = f'Биология ({course} курс)'
    elif 'психолого-педагогическое' in low_base:
        if 'психология и социальная' in low_prof:
            group_title = f'Психология и социальная педагогика ({course} курс)'
        elif profile_raw:
            clean_p = re.sub(r'[«»"]', '', profile_raw).strip()
            group_title = f'{clean_p} ({course} курс)'
        else:
            group_title = f'Психолого-педагогическое образование ({course} курс)'
    elif 'педагогическое' in low_base:
        if 'иностранный' in low_prof and 'родного языка' in low_prof:
            group_title = f'Иностранный язык и Родной язык и литература ({course} курс)'
        elif 'дошкольное' in low_prof and 'начальное' in low_prof:
            group_title = f'Дошкольное образование и Начальное образование ({course} курс)'
        elif 'физ' in low_prof and 'спорт' in low_prof:
            group_title = f'Пед. образование: Физ. культура и спорт ({course} курс)'
        elif 'физ' in low_prof:
            group_title = f'Педагогическое образование: Физ. культура ({course} курс)'
        elif 'биология' in low_prof and 'химия' in low_prof:
            group_title = f'Пед. образование: Биология и Химия ({course} курс)'
        elif profile_raw:
            clean_p = re.sub(r'[«»"]', '', profile_raw)
            clean_p = re.sub(r'\s*,\s*', ' и ', clean_p).strip()
            group_title = f'{clean_p} ({course} курс)'
        else:
            group_title = f'Педагогическое образование ({course} курс)'
    else:
        clean_label = (
            f'{base_dir}: {profile_raw}'
            if profile_raw and profile_raw.lower() not in low_base
            else base_dir
        )
        clean_label = re.sub(r'[«»"]', '', clean_label).strip()
        group_title = f'{clean_label} ({course} курс)'

    return course, group_title, practice_range


def split_sub_segments(weeks_str, type_raw):
    type_fixed = re.sub(r'Лвб', 'Лаб', type_raw, flags=re.IGNORECASE)
    cleaned = re.sub(r'поток\s*(?:с\s*)?[А-Яа-яЁёA-Za-z0-9+]+', '', type_fixed, flags=re.IGNORECASE)
    cleaned = re.sub(r'^[\s\-–—]+', '', cleaned)
    cleaned = re.sub(r'\s+', ' ', cleaned).strip(' ,;-–—')

    sub_m = re.search(
        r'^(.*?),\s*(\d+(?:\s*[-–—]\s*\d+)?(?:\s*,\s*\d+(?:\s*[-–—]\s*\d+)?)*)\s*([-–—]?\s*(?:(?:\d+|[IІVХX]+)\s*(?:п/г|подгр(?:уппа)?\.?)|ЛК|ПЗ|Лаб|ЛР|ПР|Сем).*)$',
        cleaned,
        re.IGNORECASE,
    )
    if sub_m:
        t1 = re.sub(r'^[\s\-–—]+', '', sub_m.group(1)).strip(' ,;') or 'ЛК, ПЗ'
        w2 = re.sub(r'\s+', '', sub_m.group(2)).replace('–', '-').replace('—', '-')
        t2 = re.sub(r'^[\s\-–—]+', '', sub_m.group(3)).strip(' ,;') or 'ПЗ'
        w1 = re.sub(r'\s+', '', weeks_str).replace('–', '-').replace('—', '-')
        return [(w1, t1), (w2, t2)]

    sub_m2 = re.search(
        r'^(ЛК|ПЗ|Лаб|ЛР|ПР|Сем)\s+(\d+(?:\s*[-–—]\s*\d+)?(?:\s*,\s*\d+(?:\s*[-–—]\s*\d+)?)*)\s*([-–—]?\s*(?:ЛК|ПЗ|Лаб|ЛР|ПР|Сем).*)$',
        cleaned,
        re.IGNORECASE,
    )
    if sub_m2:
        t1 = sub_m2.group(1).strip(' ,;')
        w2 = re.sub(r'\s+', '', sub_m2.group(2)).replace('–', '-').replace('—', '-')
        t2 = re.sub(r'^[\s\-–—]+', '', sub_m2.group(3)).strip(' ,;')
        w1 = re.sub(r'\s+', '', weeks_str).replace('–', '-').replace('—', '-')
        return [(w1, t1), (w2, t2)]

    w_norm = re.sub(r'\s+', '', weeks_str).replace('–', '-').replace('—', '-')
    t_norm = cleaned if cleaned else 'ЛК, ПЗ'
    return [(w_norm, t_norm)]


def _extract_subject_teacher_room(pre_text):
    """
    Разбирает фрагмент текста ДО скобки с неделями на (subject, teacher, room).
    Поддерживает оба стиля:
    - ЕМФ: "Физика ауд. 304" -> subject="Физика", teacher="", room="ауд. 304"
    - ПФ:  "ПРАКТИЧЕСКАЯ ГРАММАТИКА... канд.филол.н., доц. Исхакова Р.Р." ->
           subject="ПРАКТИЧЕСКАЯ ГРАММАТИКА...", teacher="канд.филол.н., доц. Исхакова Р.Р.", room=""
    """
    txt = re.sub(r'\s+', ' ', str(pre_text or '')).strip(' ,;-–—/')
    room_clean = ''
    teacher_clean = ''

    room_m = ROOM_RE.search(txt)
    if room_m:
        room_clean = normalize_room_str(room_m.group(0))
        txt = (txt[: room_m.start()] + ' ' + txt[room_m.end() :]).strip(' ,;-–—/')
        txt = re.sub(r'\s+', ' ', txt).strip()

    t_matches = list(TEACHER_FULL_RE.finditer(txt))
    if t_matches:
        last_t = t_matches[-1]
        if last_t.end() >= len(txt) - 2 and last_t.start() > 0:
            candidate_subj = txt[: last_t.start()].strip(' ,;-–—/')
            if len(candidate_subj) >= 2:
                teacher_clean = clean_teacher_str(last_t.group(0))
                txt = candidate_subj

    txt = re.sub(r'\bк\s+анд\.', 'канд.', txt, flags=re.IGNORECASE)
    t_matches_retry = list(TEACHER_FULL_RE.finditer(txt))
    if not teacher_clean and t_matches_retry:
        last_t = t_matches_retry[-1]
        if last_t.end() >= len(txt) - 2 and last_t.start() > 0:
            candidate_subj = txt[: last_t.start()].strip(' ,;-–—/')
            if len(candidate_subj) >= 2:
                teacher_clean = clean_teacher_str(last_t.group(0))
                txt = candidate_subj

    if txt.isupper() and len(txt) > 5:
        words = txt.split(' ')
        formatted_words = []
        for idx_w, w in enumerate(words):
            if re.match(r'^(?:[IVX]+|Ч\.\d+|ВНД|ГТО|РФ|РБ|СССР)$', w):
                formatted_words.append(w)
            elif idx_w == 0:
                formatted_words.append(w.capitalize())
            else:
                formatted_words.append(w.lower())
        txt = ' '.join(formatted_words)

    return txt.strip(' ,;-–—/'), teacher_clean, room_clean


def parse_cell_lessons(cell_text):
    """
    Универсальный парсер ячейки таблицы расписания:
    поддерживает форматы ЕМФ (скобки до преподавателя), ПФ (преподаватель до скобок),
    несколько пар в одной ячейке через перенос строки или '//', общие аудитории и преподавателей.
    """
    flat = re.sub(r'\s+', ' ', str(cell_text or '')).strip()
    if not flat or flat in ('-', '—', '–'):
        return []

    matches = list(WEEK_BLOCK_RE.finditer(flat))
    if not matches:
        return []

    segments = []
    for i, m in enumerate(matches):
        tail_start = m.end()
        tail_end = matches[i + 1].start() if i + 1 < len(matches) else len(flat)
        between = flat[tail_start:tail_end].strip()

        if i == len(matches) - 1:
            tail_text = between
            next_subj_prefix = ''
        else:
            if '//' in between:
                left_p, right_p = between.split('//', 1)
                tail_text = left_p.strip()
                next_subj_prefix = right_p.strip()
            else:
                teacher_m = list(TEACHER_FULL_RE.finditer(between))
                room_m = list(ROOM_RE.finditer(between))

                if teacher_m:
                    first_t = teacher_m[0]
                    before_t = between[: first_t.start()].strip(' ,;-–—')
                    before_t_no_room = ROOM_RE.sub('', before_t).strip(' ,;-–—')
                    if not before_t_no_room:
                        split_pos = first_t.end()
                        after_t = between[split_pos:]
                        imm_room = re.match(r'^\s*' + ROOM_RE.pattern, after_t, re.IGNORECASE)
                        if imm_room:
                            split_pos += imm_room.end()
                        tail_text = between[:split_pos].strip()
                        next_subj_prefix = between[split_pos:].strip()
                    else:
                        imm_room_start = re.match(r'^\s*' + ROOM_RE.pattern, between, re.IGNORECASE)
                        if imm_room_start:
                            tail_text = between[: imm_room_start.end()].strip()
                            next_subj_prefix = between[imm_room_start.end() :].strip()
                        else:
                            tail_text = ''
                            next_subj_prefix = between
                elif room_m:
                    split_pos = room_m[0].end()
                    tail_text = between[:split_pos].strip()
                    next_subj_prefix = between[split_pos:].strip()
                else:
                    tail_text = ''
                    next_subj_prefix = between

        segments.append({'match': m, 'tail': tail_text, 'next_subj': next_subj_prefix})

    raw_entries = []
    for i, seg in enumerate(segments):
        m = seg['match']
        pre_raw = flat[: m.start()] if i == 0 else segments[i - 1]['next_subj']
        subj, teacher_from_pre, room_from_pre = _extract_subject_teacher_room(pre_raw)

        tail = seg['tail']
        room_clean = room_from_pre
        room_m = ROOM_RE.search(tail)
        if room_m:
            room_clean = normalize_room_str(room_m.group(0))
            tail_no_room = (tail[: room_m.start()] + ' ' + tail[room_m.end() :]).strip(' ,;-–—/')
        else:
            tail_no_room = tail.strip(' ,;-–—/')

        teacher_from_tail = clean_teacher_str(tail_no_room)
        teacher_clean = teacher_from_tail or teacher_from_pre

        raw_entries.append({
            'subject': subj,
            'weeks_raw': m.group(1),
            'type_raw': m.group(2),
            'teacher': teacher_clean,
            'room': room_clean,
        })

    # Наследование аудитории и преподавателя назад (если указаны только у последней пары в ячейке)
    for i in range(len(raw_entries) - 2, -1, -1):
        if not raw_entries[i]['room'] and raw_entries[i + 1]['room']:
            raw_entries[i]['room'] = raw_entries[i + 1]['room']
        if not raw_entries[i]['teacher'] and raw_entries[i + 1]['teacher']:
            raw_entries[i]['teacher'] = raw_entries[i + 1]['teacher']

    # Наследование аудитории вперёд (если у 2-й пары в той же ячейке кабинет не дублируется)
    for i in range(1, len(raw_entries)):
        if not raw_entries[i]['room'] and raw_entries[i - 1]['room']:
            raw_entries[i]['room'] = raw_entries[i - 1]['room']

    result = []
    for ent in raw_entries:
        if not ent['subject']:
            continue
        for w_str, l_type in split_sub_segments(ent['weeks_raw'], ent['type_raw']):
            result.append((
                ent['subject'],
                l_type,
                ent['teacher'] or 'Не указан',
                ent['room'] or 'Не указана',
                w_str,
            ))
    return result


def convert_doc_to_docx_if_needed(file_path):
    """Если передан старый .doc, пытается конвертировать его в .docx через LibreOffice/soffice."""
    if not file_path.lower().endswith('.doc') or file_path.lower().endswith('.docx'):
        return file_path, None

    office_bin = shutil.which('libreoffice') or shutil.which('soffice')
    if not office_bin:
        return None, (
            '⚠️ <b>Файл в старом формате .doc, а на сервере не установлен LibreOffice для автоконвертации.</b>\n\n'
            'Открой файл в Word / Google Документах, нажми <b>«Сохранить как → Документ Word (*.docx)»</b> '
            'и отправь полученный файл <code>.docx</code> сюда.'
        )

    out_dir = tempfile.mkdtemp()
    try:
        subprocess.run(
            [office_bin, '--headless', '--convert-to', 'docx', '--outdir', out_dir, file_path],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=30,
        )
        base_name = os.path.splitext(os.path.basename(file_path))[0] + '.docx'
        converted_path = os.path.join(out_dir, base_name)
        if os.path.exists(converted_path):
            return converted_path, None
        shutil.rmtree(out_dir, ignore_errors=True)
        return None, '❌ Не удалось сконвертировать .doc в .docx.'
    except Exception as e:
        shutil.rmtree(out_dir, ignore_errors=True)
        return None, f'❌ Ошибка конвертации .doc: <code>{esc(e)}</code>'


def parse_schedule_docx_file(docx_path, default_faculty='ЕМФ'):
    """
    Читает .docx файл расписания СИ УУНиТ и возвращает:
    (records_list, practice_dict, semester_title, detected_faculty)
    """
    doc = docx.Document(docx_path)

    semester_title = ''
    detected_faculty = None
    for p in doc.paragraphs:
        t = re.sub(r'\s+', ' ', p.text or '').strip()
        if not detected_faculty:
            detected_faculty = detect_faculty_from_text(t)
        if 'полугодие' in t.lower() or 'учебного года' in t.lower():
            if not semester_title:
                semester_title = t

    faculty_code = detected_faculty or default_faculty or 'ЕМФ'
    records = []
    practice_dict = {}

    for table in doc.tables:
        grid = build_logical_grid(table)
        if len(grid) < 2:
            continue

        header_row_idx = None
        col_groups = {}
        col_default_rooms = {}
        for r_idx in range(min(4, len(grid))):
            row_tcs = grid[r_idx]
            found_any = False
            temp_cols = {}
            for c_idx, tc in enumerate(row_tcs):
                txt = get_tc_clean_text(tc)
                crs, grp_title, prac_range = parse_column_header(txt)
                if crs and grp_title:
                    temp_cols[c_idx] = (crs, grp_title)
                    found_any = True
                    if prac_range:
                        practice_dict[grp_title] = list(range(prac_range[0], prac_range[1] + 1))
            if found_any:
                header_row_idx = r_idx
                col_groups = temp_cols
                if r_idx + 1 < len(grid):
                    next_row = grid[r_idx + 1]
                    for c_idx in col_groups:
                        if c_idx < len(next_row):
                            sub_txt = get_tc_clean_text(next_row[c_idx]).strip()
                            rm_m = re.search(r'(\d+)\s*ауд', sub_txt, re.IGNORECASE)
                            if rm_m:
                                col_default_rooms[c_idx] = f'ауд. {rm_m.group(1)}'
                break

            # Поддержка двухстрочной шапки (как на ФЭП: 1-я строка — "2 курс", 2-я строка — "Направление ... 365 ауд.")
            if r_idx + 1 < len(grid):
                next_row_tcs = grid[r_idx + 1]
                found_two_row = False
                temp_cols_2 = {}
                temp_rooms_2 = {}
                for c_idx, tc in enumerate(row_tcs):
                    top_txt = get_tc_clean_text(tc)
                    bot_txt = get_tc_clean_text(next_row_tcs[c_idx]) if c_idx < len(next_row_tcs) else ''
                    combined_txt = f'{top_txt} {bot_txt}'.strip()
                    crs, grp_title, prac_range = parse_column_header(combined_txt)
                    if crs and grp_title:
                        temp_cols_2[c_idx] = (crs, grp_title)
                        found_two_row = True
                        if prac_range:
                            practice_dict[grp_title] = list(range(prac_range[0], prac_range[1] + 1))
                        rm_m = re.search(r'(?:ауд\.?\s*(\d+)|(\d+)\s*ауд)', bot_txt, re.IGNORECASE)
                        if rm_m:
                            room_num = rm_m.group(1) or rm_m.group(2)
                            temp_rooms_2[c_idx] = f'ауд. {room_num}'
                if found_two_row:
                    header_row_idx = r_idx + 1
                    col_groups = temp_cols_2
                    col_default_rooms = temp_rooms_2
                    break

        if header_row_idx is None or not col_groups:
            continue

        current_day = None
        for r_idx in range(header_row_idx + 1, len(grid)):
            row_tcs = grid[r_idx]
            if len(row_tcs) < 3:
                continue

            day_candidate = normalize_day_name(get_tc_clean_text(row_tcs[0]))
            if day_candidate:
                current_day = day_candidate

            time_slot = normalize_time_slot(get_tc_clean_text(row_tcs[1]))
            if not time_slot:
                time_slot = normalize_time_slot(get_tc_clean_text(row_tcs[0]))

            if not current_day or not time_slot:
                continue

            seen_tcs = set()
            for c_idx, (crs, grp_title) in col_groups.items():
                if c_idx >= len(row_tcs):
                    continue
                tc = row_tcs[c_idx]
                if tc is None or id(tc) in seen_tcs:
                    continue
                seen_tcs.add(id(tc))

                cell_txt = get_tc_clean_text(tc)
                lessons = parse_cell_lessons(cell_txt)
                default_room = col_default_rooms.get(c_idx, 'Не указана')
                for subj, l_type, teacher, room, weeks_str in lessons:
                    final_room = room if room != 'Не указана' else default_room
                    records.append({
                        'faculty': faculty_code,
                        'course': crs,
                        'group': grp_title,
                        'day': current_day,
                        'time': time_slot,
                        'subject': subj,
                        'type': l_type,
                        'teacher': teacher,
                        'room': final_room,
                        'weeks': weeks_str,
                    })

    return records, practice_dict, semester_title, faculty_code


def get_academic_year_label(start_dt):
    return f'{start_dt.year}/{start_dt.year + 1}'


def load_semester_meta():
    global ACADEMIC_YEAR_START_DATE, PRACTICE_WEEKS
    ACADEMIC_YEAR_START_DATE = get_default_academic_year_start()
    if os.path.exists(SEMESTER_META_PATH):
        try:
            with open(SEMESTER_META_PATH, 'r', encoding='utf-8') as f:
                meta = json.load(f)
            start_raw = meta.get('academic_year_start') or meta.get('start_date')
            if start_raw:
                dt = datetime.strptime(start_raw, '%d.%m.%Y').replace(tzinfo=LOCAL_TZ)
                ACADEMIC_YEAR_START_DATE = dt
            if isinstance(meta.get('practice_weeks'), dict):
                PRACTICE_WEEKS = {k: set(v) for k, v in meta['practice_weeks'].items()}
        except Exception as e:
            print(f'Ошибка чтения semester_meta.json: {e}')


def save_semester_meta(
    academic_start_dt,
    practice_dict,
    semester_title='',
    fac_groups=None,
    reset_done_faculties=None,
    academic_year_label=None,
):
    global ACADEMIC_YEAR_START_DATE, PRACTICE_WEEKS
    ACADEMIC_YEAR_START_DATE = academic_start_dt.replace(tzinfo=LOCAL_TZ)
    if fac_groups:
        for grp in fac_groups:
            PRACTICE_WEEKS.pop(grp, None)
    for k, v in practice_dict.items():
        PRACTICE_WEEKS[k] = set(v)
    existing_meta = {}
    if os.path.exists(SEMESTER_META_PATH):
        try:
            with open(SEMESTER_META_PATH, 'r', encoding='utf-8') as f:
                existing_meta = json.load(f)
        except Exception:
            existing_meta = {}
    yr_label = academic_year_label or get_academic_year_label(ACADEMIC_YEAR_START_DATE)
    meta = {
        'academic_year': yr_label,
        'academic_year_start': ACADEMIC_YEAR_START_DATE.strftime('%d.%m.%Y'),
        'semester_title': semester_title,
        'practice_weeks': {k: sorted(list(v)) for k, v in PRACTICE_WEEKS.items()},
        'reset_done_faculties': reset_done_faculties if reset_done_faculties is not None else existing_meta.get('reset_done_faculties', []),
        'updated_at': now_local().strftime('%d.%m.%Y %H:%M'),
    }
    atomic_write_json(SEMESTER_META_PATH, meta)


# --- ПОСТОЯННОЕ ХРАНИЛИЩЕ СОСТОЯНИЙ ПОЛЬЗОВАТЕЛЕЙ (RAM-КЭШ + SQLITE user_states) ---
STATE_TTL_SECONDS = 3600 * 12
_USER_STATES_CACHE = {}
_USER_STATES_LOCK = threading.RLock()


@contextmanager
def db_connection():
    conn = sqlite3.connect(DB_PATH, timeout=30.0, check_same_thread=False)
    try:
        conn.execute('PRAGMA busy_timeout = 30000')
        conn.execute('PRAGMA foreign_keys = ON')
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def _load_user_states_cache():
    now_ts = time.time()
    loaded = {}
    with db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('SELECT user_id, state, payload_json, updated_at FROM user_states')
        for uid, state_name, payload_json, updated_at in cursor.fetchall():
            ts = float(updated_at or 0)
            if now_ts - ts > STATE_TTL_SECONDS:
                continue
            try:
                payload = json.loads(payload_json) if payload_json else {}
            except Exception:
                payload = {}
            loaded[int(uid)] = (state_name or '', payload, ts)
    with _USER_STATES_LOCK:
        _USER_STATES_CACHE.clear()
        _USER_STATES_CACHE.update(loaded)


def set_user_state(user_id, state_name='', payload=None):
    uid = int(user_id)
    st_str = str(state_name or '')
    pl_dict = dict(payload) if payload else {}
    payload_str = json.dumps(pl_dict, ensure_ascii=False)
    now_ts = time.time()
    with _USER_STATES_LOCK:
        _USER_STATES_CACHE[uid] = (st_str, pl_dict, now_ts)
    with db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            'INSERT OR REPLACE INTO user_states (user_id, state, payload_json, updated_at) '
            'VALUES (?, ?, ?, ?)',
            (uid, st_str, payload_str, now_ts),
        )


def get_user_state(user_id):
    uid = int(user_id)
    now_ts = time.time()
    with _USER_STATES_LOCK:
        entry = _USER_STATES_CACHE.get(uid)
        if not entry:
            return '', {}
        state_name, payload, updated_at = entry
        if now_ts - updated_at > STATE_TTL_SECONDS:
            _USER_STATES_CACHE.pop(uid, None)
        else:
            return state_name, dict(payload)

    with db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('DELETE FROM user_states WHERE user_id = ?', (uid,))
    return '', {}


def clear_user_state(user_id, keep_payload=False):
    uid = int(user_id)
    if keep_payload:
        _, old_payload = get_user_state(uid)
        set_user_state(uid, '', old_payload)
    else:
        with _USER_STATES_LOCK:
            _USER_STATES_CACHE.pop(uid, None)
        with db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('DELETE FROM user_states WHERE user_id = ?', (uid,))


def cleanup_stale_states():
    now_ts = time.time()
    with _USER_STATES_LOCK:
        expired_uids = [
            uid for uid, (_, _, ts) in _USER_STATES_CACHE.items()
            if now_ts - ts > STATE_TTL_SECONDS
        ]
        for uid in expired_uids:
            _USER_STATES_CACHE.pop(uid, None)
    with db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('DELETE FROM user_states WHERE ? - updated_at > ?', (now_ts, STATE_TTL_SECONDS))
    with media_group_lock:
        stale_mgs = [
            mg_id for mg_id, ent in media_group_buffer.items()
            if now_ts - float(ent.get('_created_ts', 0)) > 120
        ]
        for mg_id in stale_mgs:
            ent = media_group_buffer.pop(mg_id, None)
            if ent and ent.get('timer'):
                try:
                    ent['timer'].cancel()
                except Exception:
                    pass
    for k in list(pending_docx_imports.keys()):
        if now_ts - pending_docx_imports[k].get('_ts', 0) > 3600:
            get_pending_docx_import(k, pop=True)
    if os.path.isdir(PENDING_DOCX_DIR):
        for fn in os.listdir(PENDING_DOCX_DIR):
            fp = os.path.join(PENDING_DOCX_DIR, fn)
            try:
                if now_ts - os.path.getmtime(fp) > 3600:
                    os.remove(fp)
            except Exception:
                pass


def cleanup_old_db_records(days_keep=30):
    cutoff = (now_local() - timedelta(days=days_keep)).date()
    with db_connection() as conn:
        cursor = conn.cursor()
        for table in ('replacements', 'cancelled_lessons'):
            cursor.execute(f'SELECT id, target_date FROM {table}')
            stale_ids = []
            for rec_id, d_str in cursor.fetchall():
                try:
                    if datetime.strptime(d_str, '%d.%m.%Y').date() < cutoff:
                        stale_ids.append((rec_id,))
                except ValueError:
                    pass
            if stale_ids:
                cursor.executemany(f'DELETE FROM {table} WHERE id = ?', stale_ids)

        cursor.execute('SELECT id, target_date, queue_id FROM homeworks')
        stale_hw_ids = []
        stale_q_ids = []
        for hw_id, d_str, q_id in cursor.fetchall():
            try:
                if datetime.strptime(d_str, '%d.%m.%Y').date() < cutoff:
                    stale_hw_ids.append((hw_id,))
                    if q_id:
                        stale_q_ids.append((q_id,))
            except ValueError:
                pass
        if stale_q_ids:
            cursor.executemany('DELETE FROM queue_slots WHERE queue_id = ?', stale_q_ids)
            cursor.executemany('DELETE FROM queues WHERE id = ?', stale_q_ids)
        if stale_hw_ids:
            cursor.executemany('DELETE FROM homeworks WHERE id = ?', stale_hw_ids)
        cursor.execute('PRAGMA optimize')
    try:
        wal_conn = sqlite3.connect(DB_PATH, timeout=30.0)
        wal_conn.execute('PRAGMA busy_timeout = 30000')
        wal_conn.execute('PRAGMA wal_checkpoint(TRUNCATE)')
        wal_conn.close()
    except Exception as e:
        print(f'Ошибка при выполнении WAL checkpoint: {e}')


# --- БАЗА ДАННЫХ И ИНИЦИАЛИЗАЦИЯ ---
def init_db():
    load_semester_meta()
    with db_connection() as conn:
        conn.execute('PRAGMA journal_mode = WAL')
        conn.execute('PRAGMA synchronous = NORMAL')
        cursor = conn.cursor()

        cursor.execute(
            'CREATE TABLE IF NOT EXISTS users ('
            'user_id INTEGER PRIMARY KEY, '
            'faculty TEXT DEFAULT "ЕМФ", '
            'course INTEGER, '
            'group_name TEXT, '
            'subscribed INTEGER DEFAULT 0, '
            'username TEXT DEFAULT "", '
            'is_starosta INTEGER DEFAULT 0, '
            'is_hw_manager INTEGER DEFAULT 0, '
            'notify_updates INTEGER DEFAULT 0, '
            'managed_group TEXT DEFAULT "")'
        )

        cursor.execute(
            'CREATE TABLE IF NOT EXISTS user_states ('
            'user_id INTEGER PRIMARY KEY, '
            'state TEXT DEFAULT "", '
            'payload_json TEXT DEFAULT "{}", '
            'updated_at REAL)'
        )

        cursor.execute("PRAGMA table_info(users)")
        existing_cols = [col[1] for col in cursor.fetchall()]
        if 'faculty' not in existing_cols:
            cursor.execute('ALTER TABLE users ADD COLUMN faculty TEXT DEFAULT "ЕМФ"')
            cursor.execute('UPDATE users SET faculty = "ЕМФ" WHERE faculty IS NULL OR faculty = ""')
        if 'username' not in existing_cols:
            cursor.execute('ALTER TABLE users ADD COLUMN username TEXT DEFAULT ""')
        if 'is_starosta' not in existing_cols:
            cursor.execute('ALTER TABLE users ADD COLUMN is_starosta INTEGER DEFAULT 0')
        if 'is_hw_manager' not in existing_cols:
            cursor.execute('ALTER TABLE users ADD COLUMN is_hw_manager INTEGER DEFAULT 0')
        if 'notify_updates' not in existing_cols:
            cursor.execute('ALTER TABLE users ADD COLUMN notify_updates INTEGER DEFAULT 0')
        if 'managed_group' not in existing_cols:
            cursor.execute('ALTER TABLE users ADD COLUMN managed_group TEXT DEFAULT ""')
            cursor.execute(
                'UPDATE users SET managed_group = group_name '
                'WHERE (is_starosta = 1 OR is_hw_manager = 1) AND (managed_group IS NULL OR managed_group = "")'
            )

        cursor.execute(
            'CREATE TABLE IF NOT EXISTS replacements ('
            'id INTEGER PRIMARY KEY AUTOINCREMENT, '
            'group_name TEXT, '
            'target_date TEXT, '
            'time_slot TEXT, '
            'subject_info TEXT, '
            'added_by INTEGER)'
        )

        cursor.execute(
            'CREATE TABLE IF NOT EXISTS cancelled_lessons ('
            'id INTEGER PRIMARY KEY AUTOINCREMENT, '
            'group_name TEXT, '
            'target_date TEXT, '
            'time_slot TEXT, '
            'cancelled_by INTEGER)'
        )

        cursor.execute(
            'CREATE TABLE IF NOT EXISTS homeworks ('
            'id INTEGER PRIMARY KEY AUTOINCREMENT, '
            'group_name TEXT, '
            'target_date TEXT, '
            'subject_name TEXT, '
            'hw_text TEXT, '
            'photo_file_id TEXT DEFAULT "", '
            'doc_file_id TEXT DEFAULT "", '
            'doc_file_name TEXT DEFAULT "", '
            'queue_id INTEGER DEFAULT NULL, '
            'added_by INTEGER, '
            'created_at TEXT)'
        )

        cursor.execute("PRAGMA table_info(homeworks)")
        hw_cols = [col[1] for col in cursor.fetchall()]
        if 'doc_file_id' not in hw_cols:
            cursor.execute('ALTER TABLE homeworks ADD COLUMN doc_file_id TEXT DEFAULT ""')
        if 'doc_file_name' not in hw_cols:
            cursor.execute('ALTER TABLE homeworks ADD COLUMN doc_file_name TEXT DEFAULT ""')
        if 'queue_id' not in hw_cols:
            cursor.execute('ALTER TABLE homeworks ADD COLUMN queue_id INTEGER DEFAULT NULL')

        cursor.execute(
            'CREATE TABLE IF NOT EXISTS queues ('
            'id INTEGER PRIMARY KEY AUTOINCREMENT, '
            'group_name TEXT, '
            'title TEXT, '
            'max_slots INTEGER DEFAULT 25, '
            'is_open INTEGER DEFAULT 1, '
            'homework_id INTEGER DEFAULT NULL, '
            'created_by INTEGER, '
            'created_at TEXT)'
        )

        cursor.execute("PRAGMA table_info(queues)")
        q_cols = [col[1] for col in cursor.fetchall()]
        if 'homework_id' not in q_cols:
            cursor.execute('ALTER TABLE queues ADD COLUMN homework_id INTEGER DEFAULT NULL')

        cursor.execute(
            'CREATE TABLE IF NOT EXISTS queue_slots ('
            'id INTEGER PRIMARY KEY AUTOINCREMENT, '
            'queue_id INTEGER, '
            'slot_num INTEGER, '
            'user_id INTEGER, '
            'user_label TEXT, '
            'joined_at TEXT)'
        )
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_queues_group ON queues(group_name, is_open)')
        cursor.execute(
            'CREATE UNIQUE INDEX IF NOT EXISTS idx_qslots_num '
            'ON queue_slots(queue_id, slot_num)'
        )
        cursor.execute('DROP INDEX IF EXISTS idx_qslots_user')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_qslots_user_nonuniq ON queue_slots(queue_id, user_id)')

        cursor.execute(
            'DELETE FROM queue_slots WHERE queue_id IN '
            '(SELECT id FROM queues WHERE homework_id IS NULL)'
        )
        cursor.execute('DELETE FROM queues WHERE homework_id IS NULL')

        cursor.execute(
            'DELETE FROM replacements WHERE id NOT IN ('
            'SELECT MAX(id) FROM replacements GROUP BY group_name, target_date, time_slot)'
        )
        cursor.execute(
            'DELETE FROM cancelled_lessons WHERE id NOT IN ('
            'SELECT MAX(id) FROM cancelled_lessons GROUP BY group_name, target_date, time_slot)'
        )

        cursor.execute('CREATE INDEX IF NOT EXISTS idx_users_group ON users(group_name)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_users_faculty ON users(faculty, course)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_users_username ON users(username)')
        cursor.execute(
            'CREATE UNIQUE INDEX IF NOT EXISTS idx_repl_unique_slot '
            'ON replacements(group_name, target_date, time_slot)'
        )
        cursor.execute(
            'CREATE UNIQUE INDEX IF NOT EXISTS idx_canc_unique_slot '
            'ON cancelled_lessons(group_name, target_date, time_slot)'
        )
        cursor.execute(
            'CREATE INDEX IF NOT EXISTS idx_hw_group_date '
            'ON homeworks(group_name, target_date)'
        )

        reload_schedule_tables(cursor, conn, force=True)
    _load_user_states_cache()


def parse_weeks(weeks_str):
    weeks = set()
    parts = str(weeks_str).replace(' ', '').split(',')
    for part in parts:
        if '-' in part:
            s, e = map(int, part.split('-'))
            weeks.update(range(s, e + 1))
        elif part.isdigit():
            weeks.add(int(part))
    return sorted(list(weeks))


def reload_schedule_tables(cursor, conn, force=False):
    global FACULTY_STRUCTURE
    if not os.path.exists(SCHEDULE_JSON_PATH):
        atomic_write_json(SCHEDULE_JSON_PATH, [])

    with open(SCHEDULE_JSON_PATH, 'rb') as fb:
        raw_bytes = fb.read()
    file_hash = hashlib.sha256(raw_bytes).hexdigest()
    schedule_data = json.loads(raw_bytes.decode('utf-8'))

    cursor.execute(
        'CREATE TABLE IF NOT EXISTS app_meta ('
        'key TEXT PRIMARY KEY, '
        'value TEXT)'
    )
    cursor.execute("SELECT value FROM app_meta WHERE key = 'schedule_json_hash'")
    hash_row = cursor.fetchone()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='schedules'")
    has_sched_tbl = bool(cursor.fetchone())
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='schedule_weeks'")
    has_weeks_tbl = bool(cursor.fetchone())

    dynamic_struct = {fac: {} for fac in FACULTY_ORDER}
    parsed_items = []
    for item in schedule_data:
        if isinstance(item, dict):
            fac = item.get('faculty') or infer_faculty_by_group(item['group'])
            crs = int(item['course'])
            grp = item['group']
            day = item['day']
            time_slot = item['time']
            subj = item['subject']
            l_type = item['type']
            teacher = item['teacher']
            room = item['room']
            weeks_str = str(item['weeks'])
        else:
            crs, grp, day, time_slot, subj, l_type, teacher, room, weeks_str = item
            fac = infer_faculty_by_group(grp)
        dynamic_struct.setdefault(fac, {}).setdefault(crs, set()).add(grp)
        parsed_items.append((fac, crs, grp, day, time_slot, subj, l_type, teacher, room, weeks_str))

    new_struct = {fac: {} for fac in FACULTY_ORDER}
    for fac, courses in dynamic_struct.items():
        new_struct[fac] = {
            crs: sorted(list(grps)) for crs, grps in sorted(courses.items())
        }
    FACULTY_STRUCTURE = new_struct

    if not force and has_sched_tbl and has_weeks_tbl and hash_row and hash_row[0] == file_hash:
        return

    cursor.execute('DROP TABLE IF EXISTS schedules')
    cursor.execute('DROP TABLE IF EXISTS schedule_weeks')

    cursor.execute(
        'CREATE TABLE schedules ('
        'id INTEGER PRIMARY KEY AUTOINCREMENT, '
        'faculty TEXT DEFAULT "ЕМФ", '
        'course_number INTEGER, '
        'group_name TEXT, '
        'day_of_week TEXT, '
        'time_slot TEXT, '
        'subject_name TEXT, '
        'lesson_type TEXT, '
        'teacher TEXT, '
        'room TEXT, '
        'weeks_raw TEXT)'
    )
    cursor.execute(
        'CREATE TABLE schedule_weeks ('
        'schedule_id INTEGER, '
        'week_number INTEGER)'
    )

    weeks_batch = []
    for row_vals in parsed_items:
        cursor.execute(
            'INSERT INTO schedules ('
            'faculty, course_number, group_name, day_of_week, time_slot, '
            'subject_name, lesson_type, teacher, room, weeks_raw'
            ') VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
            row_vals,
        )
        sched_id = cursor.lastrowid
        for w in parse_weeks(row_vals[9]):
            weeks_batch.append((sched_id, w))

    if weeks_batch:
        cursor.executemany(
            'INSERT INTO schedule_weeks (schedule_id, week_number) VALUES (?, ?)',
            weeks_batch,
        )

    cursor.execute(
        'CREATE INDEX IF NOT EXISTS idx_sched_lookup '
        'ON schedules(course_number, group_name, day_of_week, time_slot)'
    )
    cursor.execute(
        'CREATE INDEX IF NOT EXISTS idx_sched_fac_lookup '
        'ON schedules(faculty, course_number, group_name, day_of_week)'
    )
    cursor.execute(
        'CREATE INDEX IF NOT EXISTS idx_sched_day '
        'ON schedules(day_of_week, time_slot)'
    )
    cursor.execute(
        'CREATE INDEX IF NOT EXISTS idx_sched_weeks_week '
        'ON schedule_weeks(week_number, schedule_id)'
    )
    cursor.execute(
        "INSERT OR REPLACE INTO app_meta (key, value) VALUES ('schedule_json_hash', ?)",
        (file_hash,),
    )


# --- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---
def esc(text):
    return html.escape(str(text or ''))


def short_text(text, max_len=28):
    text = str(text or '')
    return text if len(text) <= max_len else text[: max_len - 1] + '…'


def parse_photo_ids(raw_photo_field):
    if not raw_photo_field:
        return []
    val = str(raw_photo_field).strip()
    if not val:
        return []
    if val.startswith('['):
        try:
            parsed = json.loads(val)
            return [str(x).strip() for x in parsed if str(x).strip()]
        except Exception:
            pass
    return [p.strip() for p in val.split(',') if p.strip()]


def serialize_photo_ids(photo_ids):
    clean = [p.strip() for p in photo_ids if p and p.strip()]
    if not clean:
        return ''
    if len(clean) == 1:
        return clean[0]
    return json.dumps(clean, ensure_ascii=False)


def send_hw_media(chat_id, photo_ids, caption_html, reply_markup=None, doc_file_id='', doc_file_name=''):
    has_doc = bool(doc_file_id)
    if not photo_ids and not has_doc:
        bot.send_message(chat_id, caption_html, reply_markup=reply_markup, parse_mode='HTML')
        return

    if has_doc and not photo_ids:
        bot.send_document(
            chat_id,
            doc_file_id,
            caption=caption_html[:1024],
            reply_markup=reply_markup,
            parse_mode='HTML',
        )
        return

    if len(photo_ids) == 1 and not has_doc:
        bot.send_photo(
            chat_id,
            photo_ids[0],
            caption=caption_html[:1024],
            reply_markup=reply_markup,
            parse_mode='HTML',
        )
        return

    if len(photo_ids) == 1 and has_doc:
        bot.send_photo(
            chat_id,
            photo_ids[0],
            caption=caption_html[:1024],
            parse_mode='HTML',
        )
        bot.send_document(
            chat_id,
            doc_file_id,
            caption=f'📎 <b>Прикреплённый файл:</b> {esc(doc_file_name or "Документ")}',
            reply_markup=reply_markup,
            parse_mode='HTML',
        )
        return

    media = []
    for idx, pid in enumerate(photo_ids):
        if idx == 0:
            media.append(InputMediaPhoto(pid, caption=caption_html[:1024], parse_mode='HTML'))
        else:
            media.append(InputMediaPhoto(pid))
    bot.send_media_group(chat_id, media)
    if has_doc:
        bot.send_document(
            chat_id,
            doc_file_id,
            caption=f'📎 <b>Прикреплённый файл:</b> {esc(doc_file_name or "Документ")}',
            reply_markup=reply_markup,
            parse_mode='HTML',
        )
    elif reply_markup:
        bot.send_message(
            chat_id,
            '⚙️ <i>Управление записью выше:</i>',
            reply_markup=reply_markup,
            parse_mode='HTML',
        )


def _direct_send_message(user_id, text, reply_markup=None, parse_mode='HTML'):
    try:
        bot.send_message(user_id, text, reply_markup=reply_markup, parse_mode=parse_mode)
        return 'ok', 0
    except ApiTelegramException as e:
        if e.error_code == 403:
            with db_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    'UPDATE users SET subscribed = 0, notify_updates = 0 WHERE user_id = ?',
                    (user_id,),
                )
            invalidate_user_cache(user_id)
            return 'blocked', 0
        elif e.error_code == 429:
            retry_after = int(getattr(e, 'result_json', {}).get('parameters', {}).get('retry_after', 2))
            return 'retry', retry_after
        return 'error', 0
    except Exception:
        return 'error', 0


def _direct_send_hw_media(user_id, photo_ids, caption_html, reply_markup=None, doc_file_id='', doc_file_name=''):
    try:
        send_hw_media(
            user_id,
            photo_ids,
            caption_html,
            reply_markup=reply_markup,
            doc_file_id=doc_file_id,
            doc_file_name=doc_file_name,
        )
        return 'ok', 0
    except ApiTelegramException as e:
        if e.error_code == 403:
            with db_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    'UPDATE users SET subscribed = 0, notify_updates = 0 WHERE user_id = ?',
                    (user_id,),
                )
            invalidate_user_cache(user_id)
            return 'blocked', 0
        elif e.error_code == 429:
            retry_after = int(getattr(e, 'result_json', {}).get('parameters', {}).get('retry_after', 2))
            return 'retry', retry_after
        return 'error', 0
    except Exception:
        return 'error', 0


def split_long_html_message(text, max_len=3900):
    if len(text) <= max_len:
        return [text]

    parts = []
    blocks = text.split('\n\n')
    current = ''
    for block in blocks:
        candidate = f'{current}\n\n{block}' if current else block
        if len(candidate) <= max_len:
            current = candidate
        else:
            if current:
                parts.append(current)
                current = ''
            if len(block) <= max_len:
                current = block
            else:
                lines = block.split('\n')
                for line in lines:
                    cand_line = f'{current}\n{line}' if current else line
                    if len(cand_line) <= max_len:
                        current = cand_line
                    else:
                        if current:
                            parts.append(current)
                        current = line[:max_len]
    if current:
        parts.append(current)
    return parts


def get_pair_num(time_slot):
    clean_slot = time_slot.strip()
    if clean_slot in PAIR_NUMBERS:
        return PAIR_NUMBERS[clean_slot]
    for slot, num in PAIR_NUMBERS.items():
        if clean_slot.startswith(slot[:5]):
            return num
    return None


def format_pair_prefix(time_slot):
    num = get_pair_num(time_slot)
    if num:
        emoji = PAIR_EMOJIS.get(num, f'{num}.')
        return f'{emoji} <b>{num} пара</b> <code>({time_slot})</code>'
    return f'⏰ <code>{esc(time_slot)}</code>'


def is_admin_user(user_obj):
    return bool(user_obj and user_obj.username and user_obj.username.lower() == ADMIN_USERNAME.lower())


_USER_DATA_CACHE = {}
_USER_USERNAME_CACHE = {}
_USER_CACHE_LOCK = threading.RLock()


def invalidate_user_cache(user_id=None):
    with _USER_CACHE_LOCK:
        if user_id is None:
            _USER_DATA_CACHE.clear()
            _USER_USERNAME_CACHE.clear()
        else:
            uid = int(user_id)
            _USER_DATA_CACHE.pop(uid, None)
            _USER_USERNAME_CACHE.pop(uid, None)


_ADMIN_COMMANDS_SYNCED_IDS = set()


def sync_admin_personal_commands(admin_user_id):
    uid = int(admin_user_id)
    if uid in _ADMIN_COMMANDS_SYNCED_IDS:
        return
    try:
        bot.set_my_commands(
            [
                BotCommand('start', 'Перезапустить бота / главное меню'),
                BotCommand('starosta', 'Выдать роль старосты (админ)'),
                BotCommand('unstarosta', 'Снять роль старосты (админ)'),
                BotCommand('stats', 'Статистика бота (админ)'),
                BotCommand('broadcast', 'Массовая рассылка объявлений (админ)'),
                BotCommand('backup', 'Скачать резервную копию БД (админ)'),
            ],
            scope=BotCommandScopeChat(uid),
        )
        _ADMIN_COMMANDS_SYNCED_IDS.add(uid)
    except Exception:
        pass


def sync_username(user_obj):
    if not user_obj:
        return
    uid = int(user_obj.id)
    uname = (user_obj.username or '').lower()
    is_admin = 1 if uname == ADMIN_USERNAME.lower() else 0
    if is_admin:
        sync_admin_personal_commands(uid)
    with _USER_CACHE_LOCK:
        cached_uname = _USER_USERNAME_CACHE.get(uid)
        if cached_uname == uname and not is_admin:
            return
        if cached_uname == uname and is_admin:
            u_data = _USER_DATA_CACHE.get(uid)
            if u_data and u_data[3] == 1 and u_data[4] == 1:
                return
    with db_connection() as conn:
        cursor = conn.cursor()
        if is_admin:
            cursor.execute(
                'UPDATE users SET username = ?, is_starosta = 1, is_hw_manager = 1 WHERE user_id = ?',
                (uname, uid),
            )
        else:
            cursor.execute(
                'UPDATE users SET username = ? WHERE user_id = ?',
                (uname, uid),
            )
    with _USER_CACHE_LOCK:
        _USER_USERNAME_CACHE[uid] = uname
        _USER_DATA_CACHE.pop(uid, None)


def get_current_week_and_day(target_date=None):
    """Сквозной подсчёт учебных недель от старта учебного года без сброса во 2-м семестре.
    Возвращает отрицательное число (или 0), если дата раньше старта нового учебного года."""
    if target_date is None:
        target_date = now_local()
    delta_days = (target_date.date() - ACADEMIC_YEAR_START_DATE.date()).days
    week_number = 0 if delta_days < 0 else (delta_days // 7) + 1
    day_name = DAYS_RUS[target_date.weekday()]
    return week_number, day_name


def get_user_data(user_id):
    uid = int(user_id)
    with _USER_CACHE_LOCK:
        if uid in _USER_DATA_CACHE:
            return _USER_DATA_CACHE[uid]
    with db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            'SELECT course, group_name, subscribed, is_starosta, is_hw_manager, notify_updates, managed_group, faculty '
            'FROM users WHERE user_id = ?',
            (uid,),
        )
        row = cursor.fetchone()
    if not row:
        return None
    fac = row[7] if row[7] else infer_faculty_by_group(row[1])
    res = (row[0], row[1], row[2], row[3], row[4], row[5], row[6], fac)
    with _USER_CACHE_LOCK:
        _USER_DATA_CACHE[uid] = res
    return res


def find_course_and_faculty_by_group(group_name):
    for fac_code, courses in FACULTY_STRUCTURE.items():
        for crs, groups in courses.items():
            if group_name in groups:
                return crs, fac_code
    return None, None


def find_course_by_group(group_name):
    crs, _ = find_course_and_faculty_by_group(group_name)
    return crs


def get_managed_context(user_obj):
    data = get_user_data(user_obj.id)
    if not data:
        return None, None
    course, group_name, _, is_st, _, _, managed_group, _ = data
    if is_admin_user(user_obj):
        return course, group_name
    if is_st == 1 and managed_group:
        target_grp = managed_group
        target_crs = find_course_by_group(target_grp) or course
        return target_crs, target_grp
    return course, group_name


def is_user_starosta(user_obj):
    if is_admin_user(user_obj):
        return True
    data = get_user_data(user_obj.id)
    return bool(data and data[3] == 1)


def can_manage_hw(user_obj):
    if is_admin_user(user_obj):
        return True
    data = get_user_data(user_obj.id)
    return bool(data and is_group_valid(data[7], data[0], data[1]))


def get_day_lessons_for_date(target_date, course, group_name):
    week_number, day_name = get_current_week_and_day(target_date)
    date_str = target_date.strftime('%d.%m.%Y')

    with db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            'SELECT s.time_slot, s.subject_name, s.lesson_type '
            'FROM schedules s '
            'JOIN schedule_weeks sw ON s.id = sw.schedule_id '
            'WHERE s.course_number = ? AND s.group_name = ? '
            'AND s.day_of_week = ? AND sw.week_number = ? '
            'ORDER BY s.time_slot ASC, s.id ASC',
            (course, group_name, day_name, week_number),
        )
        all_rows = cursor.fetchall()

        cursor.execute(
            'SELECT time_slot FROM cancelled_lessons '
            'WHERE group_name = ? AND target_date = ?',
            (group_name, date_str),
        )
        cancelled_slots = {r[0] for r in cursor.fetchall()}

        cursor.execute(
            'SELECT time_slot, subject_info FROM replacements '
            'WHERE group_name = ? AND target_date = ? '
            'ORDER BY time_slot ASC, id ASC',
            (group_name, date_str),
        )
        replacements = cursor.fetchall()

    replaced_slots = {r_time for r_time, _ in replacements}
    slot_map = {}
    subjects = []

    for time_slot, subj, l_type in all_rows:
        if time_slot in cancelled_slots or time_slot in replaced_slots:
            continue
        p_num = get_pair_num(time_slot)
        if p_num:
            slot_map.setdefault(p_num, []).append(subj)
        if subj not in subjects:
            subjects.append(subj)

    for r_time, r_subj in replacements:
        p_num = get_pair_num(r_time)
        if p_num:
            slot_map[p_num] = [f'[Замена] {r_subj}']
        if r_subj not in subjects:
            subjects.append(r_subj)

    return slot_map, subjects


def get_all_group_subjects(course, group_name):
    with db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            'SELECT DISTINCT subject_name FROM schedules '
            'WHERE course_number = ? AND group_name = ? '
            'ORDER BY subject_name ASC',
            (course, group_name),
        )
        subjects = [r[0] for r in cursor.fetchall()]
    return subjects


def get_homeworks_for_date(group_name, date_str):
    with db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            'SELECT id, subject_name, hw_text, photo_file_id, doc_file_id, doc_file_name FROM homeworks '
            'WHERE group_name = ? AND target_date = ? '
            'ORDER BY id ASC',
            (group_name, date_str),
        )
        rows = cursor.fetchall()
    return rows


def get_schedule_text(target_date, course, group_name, faculty_code=None):
    week_number, day_name = get_current_week_and_day(target_date)
    date_str = target_date.strftime('%d.%m.%Y')
    if not faculty_code:
        _, faculty_code = find_course_and_faculty_by_group(group_name)
    fac_short = get_faculty_short(faculty_code or 'ЕМФ')

    with db_connection() as conn:
        cursor = conn.cursor()
        fac_filter = faculty_code or 'ЕМФ'
        cursor.execute(
            'SELECT s.time_slot, s.subject_name, s.lesson_type, s.teacher, s.room, s.weeks_raw '
            'FROM schedules s '
            'JOIN schedule_weeks sw ON s.id = sw.schedule_id '
            'WHERE s.faculty = ? AND s.course_number = ? AND s.group_name = ? '
            'AND s.day_of_week = ? AND sw.week_number = ? '
            'ORDER BY s.time_slot ASC, s.id ASC',
            (fac_filter, course, group_name, day_name, week_number),
        )
        all_rows = cursor.fetchall()

        cursor.execute(
            'SELECT time_slot FROM cancelled_lessons '
            'WHERE group_name = ? AND target_date = ?',
            (group_name, date_str),
        )
        cancelled_slots = {r[0] for r in cursor.fetchall()}

        cursor.execute(
            'SELECT time_slot, subject_info FROM replacements '
            'WHERE group_name = ? AND target_date = ? '
            'ORDER BY time_slot ASC, id ASC',
            (group_name, date_str),
        )
        replacements = cursor.fetchall()
        replaced_slots = {r[0] for r in replacements}

        rows = [r for r in all_rows if r[0] not in cancelled_slots and r[0] not in replaced_slots]

    hw_list = get_homeworks_for_date(group_name, date_str)
    hw_by_subj = {}
    for hw_id, hw_subj, hw_text, hw_photo, hw_doc_id, hw_doc_name in hw_list:
        hw_by_subj.setdefault(hw_subj.strip().lower(), []).append(
            (hw_id, hw_subj, hw_text, hw_photo, hw_doc_id, hw_doc_name)
        )
    matched_hw_ids = set()

    week_disp = f'{week_number}-я' if week_number > 0 else 'ещё не началась'
    header = (
        f'📅 <b>{day_name.upper()}, {date_str}</b>\n'
        f'<blockquote>🏛 <b>Факультет:</b> {esc(fac_short)}\n'
        f'🎓 <b>Группа:</b> {esc(group_name)}\n'
        f'📆 <b>Учебная неделя:</b> {week_disp}</blockquote>\n\n'
    )

    body = ''
    if week_number <= 0 and not replacements:
        body = (
            f'⏳ <i>Новый учебный год стартует с <b>{ACADEMIC_YEAR_START_DATE.strftime("%d.%m.%Y")}</b> — '
            f'пока можно отдохнуть перед началом занятий!</i>\n\n'
        )
    elif week_number > 52 and not replacements:
        body = (
            '⏳ <i>Расписание на новый учебный год ещё не загружено администратором. '
            'Как только оно появится, бот уведомит об этом!</i>\n\n'
        )
    elif day_name == 'Воскресенье' and not replacements:
        body = '🎉 <i>Сегодня воскресенье — законный выходной! Набирайся сил перед новой учебной неделей.</i>\n\n'
    elif week_number in PRACTICE_WEEKS.get(group_name, set()) and not replacements:
        body = '🛠 <i>На этой неделе аудиторных пар нет — у группы проходит учебная или производственная практика.</i>\n\n'
    elif not rows and not replacements:
        if cancelled_slots:
            body = '✨ <i>Все пары на этот день были отменены — можно спокойно отдохнуть или заняться своими делами!</i>\n\n'
        else:
            body = '✨ <i>На этот день пар по расписанию нет — можно спокойно отдохнуть или заняться своими делами!</i>\n\n'
    else:
        for row in rows:
            time_slot, subject, l_type, teacher, room, weeks_raw = row
            pair_prefix = format_pair_prefix(time_slot)
            body += (
                f'{pair_prefix}\n'
                f'📘 <b>{esc(subject)}</b> <i>({esc(l_type)})</i>\n'
                f'📍 <u>{esc(room)}</u>  •  👤 <i>{esc(teacher)}</i>\n'
                f'🗓 <code>Недели: {esc(weeks_raw)}</code>\n'
            )
            subj_key = subject.strip().lower()
            attached_hws = []
            for k, items in hw_by_subj.items():
                if k == subj_key or k in subj_key or subj_key in k:
                    for item in items:
                        if item[0] not in matched_hw_ids:
                            attached_hws.append(item)
                            matched_hw_ids.add(item[0])
            for _, _, hw_text, hw_photo, hw_doc_id, hw_doc_name in attached_hws:
                p_list = parse_photo_ids(hw_photo)
                badges = []
                if p_list:
                    badges.append(f'📷 фото: {len(p_list)} шт.')
                if hw_doc_id:
                    badges.append(f'📎 {esc(hw_doc_name or "файл")}')
                badge_str = f' <i>[{", ".join(badges)}]</i>' if badges else ''
                hw_display = esc(hw_text) if hw_text else '<i>(см. вложение)</i>'
                body += f'📝 <b>ДЗ:</b> {hw_display}{badge_str}\n'
            body += '\n'

    if replacements:
        body += '⚡ <b>ИЗМЕНЕНИЯ / ЗАМЕНЫ НА ЭТОТ ДЕНЬ:</b>\n'
        for r_time, r_subj in replacements:
            r_prefix = format_pair_prefix(r_time)
            body += (
                f'🔄 {r_prefix}\n'
                f'👉 <b>{esc(r_subj)}</b>\n'
            )
            r_key = r_subj.strip().lower()
            for k, items in hw_by_subj.items():
                if k == r_key or k in r_key or r_key in k:
                    for item in items:
                        if item[0] not in matched_hw_ids:
                            matched_hw_ids.add(item[0])
                            p_list = parse_photo_ids(item[3])
                            badges = []
                            if p_list:
                                badges.append(f'📷 фото: {len(p_list)} шт.')
                            if item[4]:
                                badges.append(f'📎 {esc(item[5] or "файл")}')
                            badge_str = f' <i>[{", ".join(badges)}]</i>' if badges else ''
                            hw_display = esc(item[2]) if item[2] else '<i>(см. вложение)</i>'
                            body += f'📝 <b>ДЗ:</b> {hw_display}{badge_str}\n'
            body += '\n'

    unmatched_hws = [h for h in hw_list if h[0] not in matched_hw_ids]
    if unmatched_hws:
        body += '📚 <b>ДОМАШНИЕ ЗАДАНИЯ НА ЭТОТ ДЕНЬ:</b>\n'
        for _, hw_subj, hw_text, hw_photo, hw_doc_id, hw_doc_name in unmatched_hws:
            p_list = parse_photo_ids(hw_photo)
            badges = []
            if p_list:
                badges.append(f'📷 фото: {len(p_list)} шт.')
            if hw_doc_id:
                badges.append(f'📎 {esc(hw_doc_name or "файл")}')
            badge_str = f' <i>[{", ".join(badges)}]</i>' if badges else ''
            hw_display = esc(hw_text) if hw_text else '<i>(см. вложение)</i>'
            body += f'• <b>{esc(hw_subj)}:</b> {hw_display}{badge_str}\n'
        body += '\n'

    return (header + body).strip()


def get_week_schedule_text(target_date, course, group_name, faculty_code=None):
    week_number, _ = get_current_week_and_day(target_date)
    monday = target_date - timedelta(days=target_date.weekday())
    saturday = monday + timedelta(days=5)
    if not faculty_code:
        _, faculty_code = find_course_and_faculty_by_group(group_name)
    fac_short = get_faculty_short(faculty_code or 'ЕМФ')

    week_title = f'НА {week_number}-Ю НЕДЕЛЮ' if week_number > 0 else '(ДО СТАРТА УЧЕБНОГО ГОДА)'
    text = (
        f'🗓 <b>РАСПИСАНИЕ {week_title}</b>\n'
        f'<blockquote>📆 <b>Период:</b> {monday.strftime("%d.%m")} — {saturday.strftime("%d.%m.%Y")}\n'
        f'🏛 <b>Факультет:</b> {esc(fac_short)}  •  🎓 <b>Группа:</b> {esc(group_name)}</blockquote>\n'
    )

    if week_number <= 0:
        return text + f'\n⏳ <i>Учебный год начинается с <b>{ACADEMIC_YEAR_START_DATE.strftime("%d.%m.%Y")}</b>.</i>'
    if week_number > 52:
        return text + '\n⏳ <i>Расписание на новый учебный год ещё не загружено администратором.</i>'
    if week_number in PRACTICE_WEEKS.get(group_name, set()):
        return text + '\n🛠 <i>Всю эту неделю у вашей группы проходит практика по учебному плану.</i>'

    with db_connection() as conn:
        cursor = conn.cursor()

        has_any = False
        for day_idx in range(6):
            day_date = monday + timedelta(days=day_idx)
            day_name = DAYS_RUS[day_idx]
            date_str = day_date.strftime('%d.%m.%Y')

            fac_filter = faculty_code or 'ЕМФ'
            cursor.execute(
                'SELECT s.time_slot, s.subject_name, s.lesson_type, s.teacher, s.room '
                'FROM schedules s '
                'JOIN schedule_weeks sw ON s.id = sw.schedule_id '
                'WHERE s.faculty = ? AND s.course_number = ? AND s.group_name = ? '
                'AND s.day_of_week = ? AND sw.week_number = ? '
                'ORDER BY s.time_slot ASC, s.id ASC',
                (fac_filter, course, group_name, day_name, week_number),
            )
            all_rows = cursor.fetchall()

            cursor.execute(
                'SELECT time_slot FROM cancelled_lessons '
                'WHERE group_name = ? AND target_date = ?',
                (group_name, date_str),
            )
            cancelled_slots = {r[0] for r in cursor.fetchall()}

            cursor.execute(
                'SELECT time_slot, subject_info FROM replacements '
                'WHERE group_name = ? AND target_date = ? '
                'ORDER BY time_slot ASC, id ASC',
                (group_name, date_str),
            )
            replacements = cursor.fetchall()
            replaced_slots = {r[0] for r in replacements}
            rows = [r for r in all_rows if r[0] not in cancelled_slots and r[0] not in replaced_slots]

            cursor.execute(
                'SELECT COUNT(*) FROM homeworks WHERE group_name = ? AND target_date = ?',
                (group_name, date_str),
            )
            hw_count = cursor.fetchone()[0]

            if rows or replacements or hw_count:
                has_any = True
                hw_mark = f'  📝 ДЗ: {hw_count}' if hw_count else ''
                text += f'\n🔹 <b><u>{day_name.upper()} ({day_date.strftime("%d.%m")})</u></b>{hw_mark}\n'
                for time_slot, subject, l_type, teacher, room in rows:
                    pair_prefix = format_pair_prefix(time_slot)
                    text += (
                        f'{pair_prefix} — <b>{esc(subject)}</b> <i>({esc(l_type)})</i>\n'
                        f'   📍 <u>{esc(room)}</u> · 👤 <i>{esc(teacher)}</i>\n'
                    )
                for r_time, r_subj in replacements:
                    r_prefix = format_pair_prefix(r_time)
                    text += f'🔄 <b>[ЗАМЕНА]</b> {r_prefix} — <b>{esc(r_subj)}</b>\n'

    if not has_any:
        text += '\n✨ <i>На этой неделе учебных пар в расписании не нашлось.</i>'
    return text


def send_week_schedule_message(chat_id, target_date, course, group_name, week_offset=0, edit_msg_id=None, faculty_code=None):
    full_text = get_week_schedule_text(target_date, course, group_name, faculty_code=faculty_code)
    chunks = split_long_html_message(full_text, max_len=3900)
    markup = get_week_pagination_keyboard(week_offset)

    if len(chunks) == 1:
        if edit_msg_id:
            try:
                bot.edit_message_text(
                    chunks[0],
                    chat_id,
                    edit_msg_id,
                    reply_markup=markup,
                    parse_mode='HTML',
                )
                return
            except Exception:
                pass
        bot.send_message(chat_id, chunks[0], reply_markup=markup, parse_mode='HTML')
    else:
        if edit_msg_id:
            try:
                bot.delete_message(chat_id, edit_msg_id)
            except Exception:
                pass
        for idx, chunk in enumerate(chunks):
            kb = markup if idx == len(chunks) - 1 else None
            bot.send_message(chat_id, chunk, reply_markup=kb, parse_mode='HTML')


def get_main_keyboard(subscribed=False, is_starosta=False, is_hw_manager=False):
    markup = ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    markup.row(
        KeyboardButton('📅 Сегодня'),
        KeyboardButton('🗓 Неделя'),
    )
    markup.row(
        KeyboardButton('📚 Домашка'),
        KeyboardButton('🔍 Поиск'),
    )
    if is_starosta:
        markup.row(
            KeyboardButton('⚙️ Настройки'),
            KeyboardButton('🛠 Спецвозможности'),
        )
    else:
        markup.row(KeyboardButton('⚙️ Настройки'))
    return markup


def get_role_label(user_obj):
    if is_admin_user(user_obj):
        return '👑 Администратор'
    if is_user_starosta(user_obj):
        return '🌟 Староста'
    return '🎓 Студент'


def build_faculty_keyboard():
    markup = InlineKeyboardMarkup(row_width=1)
    for fac_code in FACULTY_ORDER:
        meta = FACULTIES_META[fac_code]
        courses = FACULTY_STRUCTURE.get(fac_code, {})
        has_data = any(bool(grps) for grps in courses.values())
        badge = '' if has_data else ' (скоро)'
        markup.add(
            InlineKeyboardButton(
                f'{meta["emoji"]} {meta["short"]} — {meta["full"]}{badge}',
                callback_data=f'sel_fac_{fac_code}',
            )
        )
    return markup


def build_course_keyboard(faculty_code):
    markup = InlineKeyboardMarkup(row_width=1)
    courses = FACULTY_STRUCTURE.get(faculty_code, {})
    for c in sorted(courses.keys()):
        if courses[c]:
            markup.add(InlineKeyboardButton(f'🎓 {c} курс', callback_data=f'sel_course_{faculty_code}_{c}'))
    markup.add(InlineKeyboardButton('🔙 Сменить факультет', callback_data='back_to_faculties'))
    return markup


def open_settings_menu(chat_id, user_obj, message_id=None):
    user = get_user_data(user_obj.id)
    if not user or not is_group_valid(user[7], user[0], user[1]):
        bot.send_message(
            chat_id,
            '👇 Сначала выбери свой <b>факультет</b>:',
            reply_markup=build_faculty_keyboard(),
            parse_mode='HTML',
        )
        return

    course, group_name, subscribed, is_st, is_hw, notify_updates, managed_group, fac_code = user
    fac_short = get_faculty_short(fac_code)
    role_text = get_role_label(user_obj)
    sub_status = 'включена (07:00 и 18:00)' if subscribed else 'выключена'
    sub_btn_text = (
        '🔔 Авторассылка (07:00 и 18:00): ВКЛ'
        if subscribed
        else '🔕 Авторассылка (07:00 и 18:00): ВЫКЛ'
    )
    notify_status = 'включены' if notify_updates else 'выключены'
    notify_btn_text = (
        '🔔 Уведомления (пары и ДЗ): ВКЛ'
        if notify_updates
        else '🔕 Уведомления (пары и ДЗ): ВЫКЛ'
    )

    managed_line = ''
    if is_st == 1 and not is_admin_user(user_obj) and managed_group:
        managed_line = f'\n🛡 <b>Управляемая группа:</b> {esc(managed_group)}'

    text = (
        '⚙️ <b>НАСТРОЙКИ И ПРОФИЛЬ</b>\n\n'
        f'<blockquote>🏛 <b>Факультет:</b> {esc(fac_short)} ({esc(get_faculty_full(fac_code))})\n'
        f'🎓 <b>Курс:</b> {course}\n'
        f'📚 <b>Группа просмотра:</b> {esc(group_name)}'
        f'{managed_line}\n'
        f'👤 <b>Статус:</b> {role_text}\n'
        f'🗓 <b>Авторассылка (с ДЗ):</b> {sub_status}\n'
        f'⚡ <b>Уведомления (пары и ДЗ):</b> {notify_status}</blockquote>\n\n'
        '<i>Выбери нужный пункт ниже:</i>'
    )

    markup = InlineKeyboardMarkup(row_width=1)
    markup.add(
        InlineKeyboardButton(sub_btn_text, callback_data='settings_toggle_sub'),
        InlineKeyboardButton(notify_btn_text, callback_data='settings_toggle_notify'),
        InlineKeyboardButton('🏛 Сменить факультет / группу', callback_data='settings_change_group'),
        InlineKeyboardButton('🔄 Обновить главное меню', callback_data='settings_restart'),
    )

    if message_id:
        try:
            bot.edit_message_text(text, chat_id, message_id, reply_markup=markup, parse_mode='HTML')
            return
        except Exception:
            pass
    bot.send_message(chat_id, text, reply_markup=markup, parse_mode='HTML')


def open_special_menu(chat_id, user_obj, message_id=None):
    starosta_flag = is_user_starosta(user_obj)
    admin_flag = is_admin_user(user_obj)

    if not (starosta_flag or admin_flag):
        bot.send_message(
            chat_id,
            '⛔ Этот раздел доступен только старостам и администраторам. Добавить домашнее задание можно в разделе <b>«📚 Домашка»</b>.',
            parse_mode='HTML',
        )
        return

    user = get_user_data(user_obj.id)
    has_valid_group = bool(user and is_group_valid(user[7], user[0], user[1]))

    if not admin_flag and not has_valid_group:
        bot.send_message(chat_id, '⚠️ Сначала выбери свой факультет и группу через /start.')
        return

    _, target_group = get_managed_context(user_obj)
    group_display = esc(target_group) if (target_group and has_valid_group) else 'Не выбрана (режим админа)'
    role_text = get_role_label(user_obj)
    text = (
        '🛠 <b>СПЕЦВОЗМОЖНОСТИ</b>\n\n'
        f'<blockquote>👤 <b>Твой статус:</b> {role_text}\n'
        f'🛡 <b>Управляемая группа:</b> {group_display}</blockquote>\n\n'
        '<i>Выбери инструмент управления:</i>'
    )

    markup = InlineKeyboardMarkup(row_width=2)
    if starosta_flag and has_valid_group:
        markup.row(
            InlineKeyboardButton('➕ Замена', callback_data='spec_add_repl'),
            InlineKeyboardButton('🗑 Отмена пары', callback_data='spec_del_pair'),
        )
        markup.row(
            InlineKeyboardButton('📋 Активные замены / отмены', callback_data='spec_active_changes')
        )
        markup.row(
            InlineKeyboardButton('📝 Добавить ДЗ (с вопросами)', callback_data='spec_add_hw')
        )

    if admin_flag:
        markup.row(
            InlineKeyboardButton('📊 Статистика', callback_data='spec_stats'),
            InlineKeyboardButton('📣 Рассылка', callback_data='spec_broadcast'),
        )
        markup.row(
            InlineKeyboardButton('📥 Обновить расписание (.docx/.doc)', callback_data='spec_upload_docx')
        )
        markup.row(
            InlineKeyboardButton('💾 Бэкап БД сейчас', callback_data='spec_backup_now')
        )

    if message_id:
        try:
            bot.edit_message_text(text, chat_id, message_id, reply_markup=markup, parse_mode='HTML')
            return
        except Exception:
            pass
    bot.send_message(chat_id, text, reply_markup=markup, parse_mode='HTML')


def get_pagination_keyboard(offset=0, group_name=None):
    markup = InlineKeyboardMarkup()
    markup.row(
        InlineKeyboardButton('⬅️', callback_data=f'date_{offset - 1}'),
        InlineKeyboardButton('📍 Сегодня', callback_data='date_0'),
        InlineKeyboardButton('➡️', callback_data=f'date_{offset + 1}'),
    )
    if group_name:
        target_date_str = (now_local() + timedelta(days=offset)).strftime('%d.%m.%Y')
        hws = get_homeworks_for_date(group_name, target_date_str)
        if hws:
            has_media = any(parse_photo_ids(h[3]) or h[4] for h in hws)
            btn_label = (
                f'📎 ДЗ с вложениями ({len(hws)})'
                if has_media
                else f'📚 ДЗ на этот день ({len(hws)})'
            )
            markup.row(
                InlineKeyboardButton(
                    btn_label,
                    callback_data=f'hw_showdate_{target_date_str}',
                )
            )
    return markup


def get_week_pagination_keyboard(week_offset=0):
    markup = InlineKeyboardMarkup()
    markup.row(
        InlineKeyboardButton('⬅️ Пред. неделя', callback_data=f'week_{week_offset - 1}'),
        InlineKeyboardButton('📍 Текущая', callback_data='week_0'),
        InlineKeyboardButton('След. неделя ➡️', callback_data=f'week_{week_offset + 1}'),
    )
    return markup


def start_onboarding(message):
    sync_username(message.from_user)
    clear_user_state(message.chat.id)
    bot.send_message(
        message.chat.id,
        '👋 <b>Привет!</b>\n'
        'Я твой цифровой помощник по расписанию <i>Сибайского института (филиала) УУНиТ</i>.\n'
        f'<i>by @{ADMIN_USERNAME}</i>\n\n'
        '👇 Для начала выбери свой <b>факультет</b>:',
        reply_markup=build_faculty_keyboard(),
        parse_mode='HTML',
    )


# --- ОБРАБОТЧИК ИМПОРТА РАСПИСАНИЯ ИЗ .DOCX / .DOC (ДЛЯ АДМИНА) ---
def apply_parsed_docx_schedule(chat_id, admin_user_id, import_token, year_start_dt, faculty_code=None):
    payload = get_pending_docx_import(import_token, pop=True)
    if not payload:
        bot.send_message(chat_id, '⚠️ Сессия импорта устарела. Отправь файл расписания заново.')
        return

    target_fac = faculty_code or payload.get('faculty_code') or 'ЕМФ'
    records = payload['records']
    for r in records:
        r['faculty'] = target_fac
    practice_dict = payload['practice_dict']
    semester_title = payload['semester_title']
    clear_old_changes = payload.get('clear_old_changes', True)

    send_db_backup(manual_chat_id=chat_id)

    # Обновляем в schedule.json только записи выбранного факультета, сохраняя остальные факультеты
    existing_records = []
    if os.path.exists(SCHEDULE_JSON_PATH):
        try:
            with open(SCHEDULE_JSON_PATH, 'r', encoding='utf-8') as f:
                existing_records = json.load(f)
        except Exception:
            existing_records = []

    kept_other_faculties = [
        r for r in existing_records
        if (r.get('faculty') or infer_faculty_by_group(r.get('group', ''))) != target_fac
    ]
    merged_records = kept_other_faculties + records

    existing_meta = {}
    if os.path.exists(SEMESTER_META_PATH):
        try:
            with open(SEMESTER_META_PATH, 'r', encoding='utf-8') as f:
                existing_meta = json.load(f)
        except Exception:
            existing_meta = {}

    new_year_label = get_academic_year_label(year_start_dt)
    saved_year_label = existing_meta.get('academic_year') or get_academic_year_label(ACADEMIC_YEAR_START_DATE)

    if new_year_label != saved_year_label:
        reset_done_facs = []
    else:
        reset_done_facs = list(existing_meta.get('reset_done_faculties', []))

    # Сбрасываем курс у студентов данного факультета только один раз при переходе на новый учебный год
    is_new_academic_year = (new_year_label != saved_year_label) or (
        bool(reset_done_facs) and target_fac not in reset_done_facs
    )
    if target_fac not in reset_done_facs:
        reset_done_facs.append(target_fac)

    fac_groups = {r['group'] for r in records}

    atomic_write_json(SCHEDULE_JSON_PATH, merged_records)

    save_semester_meta(
        year_start_dt,
        practice_dict,
        semester_title,
        fac_groups=fac_groups,
        reset_done_faculties=reset_done_facs,
        academic_year_label=new_year_label,
    )

    reset_users_ids = []
    with db_connection() as conn:
        cursor = conn.cursor()
        reload_schedule_tables(cursor, conn)

        if clear_old_changes:
            for grp in fac_groups:
                cursor.execute('DELETE FROM replacements WHERE group_name = ?', (grp,))
                cursor.execute('DELETE FROM cancelled_lessons WHERE group_name = ?', (grp,))

        # При начале нового учебного года (обнулении счётчика недель) сбрасываем курс и группу ТОЛЬКО у студентов загружаемого факультета,
        # чтобы студенты других факультетов не выбирали старые группы до загрузки их расписания.
        if is_new_academic_year:
            cursor.execute(
                'SELECT user_id FROM users WHERE faculty = ? AND (group_name != "" OR course > 0)',
                (target_fac,),
            )
            reset_users_ids = [r[0] for r in cursor.fetchall()]
            cursor.execute(
                'UPDATE users SET course = 0, group_name = "", managed_group = "" WHERE faculty = ?',
                (target_fac,),
            )
            for u_id_reset in reset_users_ids:
                cursor.execute('DELETE FROM user_states WHERE user_id = ?', (u_id_reset,))
        else:
            # Если в течение года переименовали/удалили группу в файле, просим затронутых студентов перевыбрать направление
            cursor.execute(
                'SELECT user_id, group_name FROM users WHERE faculty = ? AND group_name != ""',
                (target_fac,),
            )
            orphaned_ids = [uid for uid, grp in cursor.fetchall() if grp not in fac_groups]
            if orphaned_ids:
                for uid_orph in orphaned_ids:
                    cursor.execute(
                        'UPDATE users SET group_name = "", managed_group = "" WHERE user_id = ?',
                        (uid_orph,),
                    )
                    enqueue_message(
                        uid_orph,
                        '🔄 <b>Обновился список направлений твоего факультета!</b>\n\n'
                        'Пожалуйста, выбери свою группу заново, чтобы продолжить получать расписание:',
                        reply_markup=build_faculty_keyboard(),
                        parse_mode='HTML',
                    )
                invalidate_user_cache()

    if is_new_academic_year and reset_users_ids:
        with _USER_STATES_LOCK:
            for u_id_reset in reset_users_ids:
                _USER_STATES_CACHE.pop(int(u_id_reset), None)
        invalidate_user_cache()

    if is_new_academic_year and reset_users_ids:
        fac_short_lbl = get_faculty_short(target_fac)
        for u_id in reset_users_ids:
            enqueue_message(
                u_id,
                f'🎓 <b>С НОВЫМ УЧЕБНЫМ ГОДОМ ({year_start_dt.year}/{year_start_dt.year + 1})!</b>\n\n'
                f'Загружено новое расписание для факультета <b>{esc(fac_short_lbl)}</b>, и счётчик учебных недель начался заново с <b>1-й недели</b>.\n'
                f'Чтобы продолжить получать расписание и ДЗ, пожалуйста, заново выбери свой <b>факультет, новый курс и направление</b>:',
                reply_markup=build_faculty_keyboard(),
                parse_mode='HTML',
            )

    cur_week, _ = get_current_week_and_day()
    fac_courses = FACULTY_STRUCTURE.get(target_fac, {})
    groups_summary = '\n'.join(
        f'• <b>{crs} курс:</b> {", ".join(esc(g) for g in grps)}'
        for crs, grps in sorted(fac_courses.items())
    )

    reset_note = (
        f'\n🔄 <b>Новый учебный год:</b> у {len(reset_users_ids)} пользователей сброшен курс — им отправлено меню выбора нового курса!'
        if is_new_academic_year
        else ''
    )

    bot.send_message(
        chat_id,
        f'🎉 <b>РАСПИСАНИЕ УСПЕШНО ОБНОВЛЕНО!</b>\n\n'
        f'🏛 <b>Факультет:</b> {esc(get_faculty_short(target_fac))} ({esc(get_faculty_full(target_fac))})\n'
        f'📅 <b>Точка отсчёта 1-й недели года:</b> {year_start_dt.strftime("%d.%m.%Y")} (сейчас идёт <b>{cur_week}-я</b> сквозная неделя)\n'
        f'📚 <b>Загружено пар для факультета:</b> {len(records)} (всего в базе: {len(merged_records)})\n'
        f'🧹 <b>Старые замены и отмены факультета:</b> {"очищены" if clear_old_changes else "сохранены"}'
        f'{reset_note}\n\n'
        f'🎓 <b>Активные направления ({esc(get_faculty_short(target_fac))}):</b>\n{groups_summary}\n\n'
        f'<i>Новое расписание уже работает для всех студентов и сохранено в schedule.json!</i>',
        parse_mode='HTML',
    )


@bot.message_handler(
    content_types=['document'],
    func=lambda m: bool(
        m.document
        and m.document.file_name
        and m.document.file_name.lower().endswith(('.docx', '.doc'))
        and get_user_state(m.from_user.id)[0] != 'awaiting_hw_content'
    ),
)
def handle_schedule_docx_upload(message):
    sync_username(message.from_user)
    if not is_admin_user(message.from_user):
        return

    fname = message.document.file_name
    wait_msg = bot.send_message(message.chat.id, '⏳ Скачиваю и анализирую файл расписания...')

    tmp_raw_path = None
    tmp_docx_path = None
    try:
        file_info = bot.get_file(message.document.file_id)
        downloaded = bot.download_file(file_info.file_path)

        suffix = '.docx' if fname.lower().endswith('.docx') else '.doc'
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(downloaded)
            tmp_raw_path = tmp.name

        converted_path, conv_err = convert_doc_to_docx_if_needed(tmp_raw_path)
        if conv_err:
            bot.edit_message_text(conv_err, message.chat.id, wait_msg.message_id, parse_mode='HTML')
            return
        tmp_docx_path = converted_path

        fac_from_fname = detect_faculty_from_text(fname) or 'ЕМФ'
        records, practice_dict, semester_title, detected_fac = parse_schedule_docx_file(
            tmp_docx_path, default_faculty=fac_from_fname
        )
    except Exception as e:
        bot.edit_message_text(
            f'❌ <b>Ошибка при чтении файла расписания:</b>\n<code>{esc(e)}</code>',
            message.chat.id,
            wait_msg.message_id,
            parse_mode='HTML',
        )
        return
    finally:
        for p_clean in {tmp_raw_path, tmp_docx_path}:
            if p_clean and os.path.exists(p_clean):
                try:
                    os.remove(p_clean)
                    parent_dir = os.path.dirname(p_clean)
                    if parent_dir != tempfile.gettempdir() and os.path.isdir(parent_dir) and not os.listdir(parent_dir):
                        os.rmdir(parent_dir)
                except Exception:
                    pass

    if not records:
        bot.edit_message_text(
            '⚠️ Не удалось найти таблицы с расписанием в этом файле.\n'
            'Убедись, что шапки столбцов содержат номер курса (например, <i>«3 курс Направление ...»</i>), '
            'а время пар указано в формате <code>08:00 - 09:20</code>.',
            message.chat.id,
            wait_msg.message_id,
            parse_mode='HTML',
        )
        return

    by_group = {}
    for r in records:
        by_group[(r['course'], r['group'])] = by_group.get((r['course'], r['group']), 0) + 1

    preview_lines = []
    for (crs, grp), cnt in sorted(by_group.items()):
        prac_info = ''
        if grp in practice_dict and practice_dict[grp]:
            pw = practice_dict[grp]
            prac_info = f' <i>(практика: {min(pw)}–{max(pw)} нед.)</i>'
        preview_lines.append(f'• {esc(grp)}: <b>{cnt}</b> записей{prac_info}')

    all_parsed_weeks = []
    for r in records:
        all_parsed_weeks.extend(parse_weeks(r.get('weeks', '')))
    max_parsed_week = max(all_parsed_weeks) if all_parsed_weeks else 0

    now_dt = now_local()
    spring_reset_detected = False
    if now_dt.month in (1, 2, 3, 4, 5) and 0 < max_parsed_week <= 22:
        # В весеннем семестре в файле указаны недели 1..18 (сброс нумерации с начала весеннего семестра)
        feb_1 = datetime(now_dt.year, 2, 1, tzinfo=LOCAL_TZ)
        days_to_mon = (7 - feb_1.weekday()) % 7
        suggested_start_dt = feb_1 + timedelta(days=days_to_mon)
        spring_reset_detected = True
    else:
        title_year_start = extract_academic_start_from_title(semester_title)
        if title_year_start is not None:
            suggested_start_dt = title_year_start
        else:
            auto_year_start = get_default_academic_year_start()
            suggested_start_dt = (
                auto_year_start
                if auto_year_start.date() > ACADEMIC_YEAR_START_DATE.date()
                else ACADEMIC_YEAR_START_DATE
            )
    delta_preview = (now_dt.date() - suggested_start_dt.date()).days
    cur_week = 1 if delta_preview < 0 else (delta_preview // 7) + 1
    year_start_str = suggested_start_dt.strftime('%d.%m.%Y')

    import_token = f'{message.from_user.id}_{int(time.time())}'
    save_pending_docx_import(import_token, {
        '_ts': time.time(),
        'records': records,
        'practice_dict': practice_dict,
        'semester_title': semester_title,
        'faculty_code': detected_fac,
        'year_start_dt': suggested_start_dt,
        'clear_old_changes': True,
    })

    title_line = f'📄 <b>Заголовок:</b> <i>{esc(semester_title)}</i>\n' if semester_title else ''
    text = (
        f'✅ <b>Файл расписания успешно распознан!</b>\n\n'
        f'{title_line}'
        f'🏛 <b>Определён факультет:</b> {esc(get_faculty_short(detected_fac))} ({esc(get_faculty_full(detected_fac))})\n'
        f'📊 <b>Всего строк расписания:</b> {len(records)}\n\n'
        f'🎓 <b>Распознанные группы и курсы:</b>\n'
        + '\n'.join(preview_lines)
        + (
            f'\n\n⚠️ <b>Обнаружен сброс недель весеннего семестра (1–{max_parsed_week} нед.)!</b>\n'
            f'📆 <b>Точка отсчёта 1-й недели:</b> <code>{year_start_str}</code> (сейчас <b>{cur_week}-я</b> неделя)\n\n'
            if spring_reset_detected
            else f'\n\n📆 <b>Счётчик учебных недель:</b> от <code>{year_start_str}</code> (сейчас <b>{cur_week}-я</b> неделя)\n\n'
        )
        + f'Проверь факультет и примени обновление:'
    )

    markup = InlineKeyboardMarkup(row_width=2)
    fac_btns = []
    for f_code in FACULTY_ORDER:
        mark = '✅ ' if f_code == detected_fac else ''
        fac_btns.append(
            InlineKeyboardButton(
                f'{mark}{get_faculty_short(f_code)}',
                callback_data=f'docx_setfac_{f_code}_{import_token}',
            )
        )
    markup.add(*fac_btns)
    markup.add(
        InlineKeyboardButton(
            f'🚀 Применить ({get_faculty_short(detected_fac)}, старт {year_start_str})',
            callback_data=f'docx_apply_auto_{import_token}',
        )
    )
    markup.add(
        InlineKeyboardButton(
            '📅 Изменить дату 1-й недели учебного года',
            callback_data=f'docx_apply_manual_{import_token}',
        )
    )
    markup.add(InlineKeyboardButton('❌ Отмена', callback_data=f'docx_cancel_{import_token}'))

    bot.edit_message_text(
        text,
        message.chat.id,
        wait_msg.message_id,
        reply_markup=markup,
        parse_mode='HTML',
    )


def process_docx_custom_start_date(message, payload):
    import_token = (payload or {}).get('import_token', '')
    clear_user_state(message.chat.id)
    if not message.text or message.text.strip().lower() in ['отмена', '/cancel', '/start']:
        get_pending_docx_import(import_token, pop=True)
        bot.send_message(message.chat.id, '❌ Импорт расписания отменён.')
        return

    raw = message.text.strip()
    try:
        dt = datetime.strptime(raw, '%d.%m.%Y').replace(tzinfo=LOCAL_TZ)
    except ValueError:
        set_user_state(message.chat.id, 'awaiting_docx_date', {'import_token': import_token})
        bot.send_message(
            message.chat.id,
            '⚠️ Неверный формат даты. Введи дату понедельника 1-й недели учебного года в формате <code>ДД.ММ.ГГГГ</code> (например, <code>31.08.2026</code>):',
            parse_mode='HTML',
        )
        return

    apply_parsed_docx_schedule(message.chat.id, message.from_user.id, import_token, dt)


# --- УМНЫЙ ПОИСК ПО ПРЕПОДАВАТЕЛЮ И СВОБОДНЫМ АУДИТОРИЯМ ---
TEACHER_SURNAME_INIT_RE = re.compile(
    r'([А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?)\s+([А-ЯЁ])\.\s*([А-ЯЁ])\.'
)

TEACHER_RANK_PREFIX_RE = re.compile(
    r'^(?:(?:канд|к|д-р|д)\.?[а-яё\-.\s]*?н\.?(?=[\s,.-]|$)|доц\.?(?=[\s,.-]|$)|доцент(?=[\s,.-]|$)|проф\.?(?=[\s,.-]|$)|профессор(?=[\s,.-]|$)|ст\.?\s*преп\.?(?=[\s,.-]|$)|ст\.?\s*пр\.?(?=[\s,.-]|$)|старший\s+преподаватель(?=[\s,.-]|$)|преп\.?(?=[\s,.-]|$)|преподаватель(?=[\s,.-]|$)|асс\.?(?=[\s,.-]|$)|ассистент(?=[\s,.-]|$))[\s,.-]*',
    re.IGNORECASE,
)

TEACHER_RANK_SUFFIX_RE = re.compile(
    r'(?:[,;\s]+(?:канд|к|д-р|д)\.?[а-яё\-.\s]*?н\.?|[,;\s]+(?:доц|доцент|проф|профессор|ст\.?\s*преп|ст\.?\s*пр|преп|преподаватель|асс|ассистент)\.?\s*)+$',
    re.IGNORECASE,
)


def _format_surname_part(surname_raw):
    s = re.sub(r'[^а-яёА-ЯЁ-]', '', str(surname_raw or '')).strip('-')
    if not s:
        return ''
    return '-'.join(part.capitalize() for part in s.split('-') if part)


def extract_teacher_surname_initials(raw_teacher):
    txt = re.sub(r'\s+', ' ', str(raw_teacher or '')).strip()
    if not txt or txt.lower() in ('не указан', '-', '—', 'нет', 'вакансия'):
        return ''

    if '\n' in str(raw_teacher or '') or '/' in txt or ';' in txt:
        parts_split = re.split(r'[\n/;]+', str(raw_teacher) if '\n' in str(raw_teacher or '') else txt)
        norm_parts = []
        seen_low = set()
        for p in parts_split:
            sub_norm = extract_teacher_surname_initials(p)
            if sub_norm and sub_norm.lower() not in seen_low:
                seen_low.add(sub_norm.lower())
                norm_parts.append(sub_norm)
        if norm_parts:
            return ', '.join(norm_parts)

    m = TEACHER_SURNAME_INIT_RE.search(txt)
    if m:
        surname, i1, i2 = m.groups()
        clean_surname = _format_surname_part(surname)
        if clean_surname:
            return f'{clean_surname} {i1.upper()}.{i2.upper()}.'

    cleaned = txt
    while True:
        next_c = TEACHER_RANK_PREFIX_RE.sub('', cleaned).strip(' ,;-–—/.')
        next_c = TEACHER_RANK_SUFFIX_RE.sub('', next_c).strip(' ,;-–—/.')
        if next_c == cleaned:
            break
        cleaned = next_c

    m_no_dot = re.search(
        r'([А-ЯЁа-яё]+(?:-[А-ЯЁа-яё]+)?)\s+([А-ЯЁа-яё])\.?\s*([А-ЯЁа-яё])\.?$',
        cleaned,
    )
    if m_no_dot:
        s_part, i1, i2 = m_no_dot.groups()
        f_sur = _format_surname_part(s_part)
        if f_sur and len(f_sur) >= 2:
            return f'{f_sur} {i1.upper()}.{i2.upper()}.'

    if not cleaned:
        cleaned = txt

    words = [w.strip(' ,;-–—/.()') for w in cleaned.split() if w.strip(' ,;-–—/.()')]
    if len(words) == 1 and re.fullmatch(r'[а-яёА-ЯЁ-]+', words[0]):
        return _format_surname_part(words[0])
    if words and re.fullmatch(r'[а-яёА-ЯЁ-]+', words[0]) and len(words[0]) >= 3:
        if len(words) >= 2 and re.fullmatch(r'[а-яёА-ЯЁ]\.?$', words[1]):
            init1 = words[1][0].upper()
            init2 = words[2][0].upper() if len(words) >= 3 and re.fullmatch(r'[а-яёА-ЯЁ]\.?$', words[2]) else ''
            f_sur = _format_surname_part(words[0])
            if init2:
                return f'{f_sur} {init1}.{init2}.'
            return f'{f_sur} {init1}.'
        return _format_surname_part(words[0])

    return cleaned


def _teacher_surname_only(teacher_norm):
    """Возвращает только фамилию в нижнем регистре из нормализованной строки вида 'Суюндуков Я.Т.' или 'Бускунова'."""
    first = str(teacher_norm or '').split(',')[0].strip()
    if not first:
        return ''
    return first.split()[0].lower()


def get_all_unique_teachers():
    with db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('SELECT DISTINCT teacher FROM schedules')
        raw_teachers = [r[0] for r in cursor.fetchall() if r[0]]

    by_lower = {}
    for t in raw_teachers:
        norm = extract_teacher_surname_initials(t)
        if not norm:
            continue
        for single in norm.split(', '):
            s_clean = single.strip()
            if not s_clean:
                continue
            low_key = s_clean.lower()
            if low_key not in by_lower:
                by_lower[low_key] = s_clean

    all_vals = list(by_lower.values())
    surnames_with_initials = set()
    for val in all_vals:
        if ' ' in val and '.' in val:
            surnames_with_initials.add(val.split()[0].lower())

    deduped = []
    for val in sorted(all_vals):
        low_v = val.lower()
        # Не удаляем фамилию без инициалов, если по этой фамилии есть ещё и запись без инициалов,
        # но если в базе есть "Суюндуков Я.Т." и "Суюндуков И.В.", оба останутся.
        if low_v in surnames_with_initials and ' ' not in val:
            continue
        deduped.append(val)

    return sorted(deduped)


def open_search_menu(chat_id, message_id=None):
    clear_user_state(chat_id, keep_payload=True)
    markup = InlineKeyboardMarkup(row_width=1)
    markup.add(
        InlineKeyboardButton('👨‍🏫 Где преподаватель?', callback_data='search_teacher'),
        InlineKeyboardButton('🚪 Свободные аудитории', callback_data='search_rooms'),
    )
    text = (
        '🔍 <b>УМНЫЙ ПОИСК ПО ИНСТИТУТУ</b>\n\n'
        '• <b>Где преподаватель?</b> — поиск по всем факультетам на <i>сегодня</i>, <i>завтра</i> или <i>всю учебную неделю</i> с автодополнением фамилий.\n'
        '• <b>Свободные аудитории</b> — проверка пустых кабинетов на выбранный день и пару.'
    )
    if message_id:
        try:
            bot.edit_message_text(text, chat_id, message_id, reply_markup=markup, parse_mode='HTML')
            return
        except Exception:
            pass
    bot.send_message(chat_id, text, reply_markup=markup, parse_mode='HTML')


def show_teacher_period_choice(chat_id, msg_id=None):
    today = now_local()
    tomorrow = today + timedelta(days=1)
    markup = InlineKeyboardMarkup(row_width=2)
    markup.row(
        InlineKeyboardButton(f'📍 Сегодня ({today.strftime("%d.%m")})', callback_data='tsearch_period_today'),
        InlineKeyboardButton(f'⏭ Завтра ({tomorrow.strftime("%d.%m")})', callback_data='tsearch_period_tomorrow'),
    )
    markup.row(
        InlineKeyboardButton('🗓 На всю текущую неделю', callback_data='tsearch_period_week'),
    )
    markup.row(
        InlineKeyboardButton('🔙 Назад', callback_data='search_back'),
    )
    text = (
        '👨‍🏫 <b>ПОИСК ПРЕПОДАВАТЕЛЯ</b>\n\n'
        '<b>Шаг 1 из 2.</b> Выбери период, за который нужно найти пары преподавателя:'
    )
    if msg_id:
        try:
            bot.edit_message_text(text, chat_id, msg_id, reply_markup=markup, parse_mode='HTML')
            return
        except Exception:
            pass
    bot.send_message(chat_id, text, reply_markup=markup, parse_mode='HTML')


def execute_teacher_search_for_period(chat_id, query, period):
    today = now_local()
    q_lower = query.strip().lower()

    if period == 'today':
        target_dates = [today]
        period_title = f'Сегодня ({DAYS_RUS[today.weekday()]}, {today.strftime("%d.%m.%Y")})'
    elif period == 'tomorrow':
        tmr = today + timedelta(days=1)
        target_dates = [tmr]
        period_title = f'Завтра ({DAYS_RUS[tmr.weekday()]}, {tmr.strftime("%d.%m.%Y")})'
    else:
        monday = today - timedelta(days=today.weekday())
        target_dates = [monday + timedelta(days=i) for i in range(6)]
        w_num, _ = get_current_week_and_day(today)
        period_title = f'Вся {w_num}-я неделя ({monday.strftime("%d.%m")} — {target_dates[-1].strftime("%d.%m")})'

    all_day_results = []
    with db_connection() as conn:
        cursor = conn.cursor()
        for dt_obj in target_dates:
            week_num, day_name = get_current_week_and_day(dt_obj)
            date_str = dt_obj.strftime('%d.%m.%Y')

            cursor.execute(
                'SELECT s.time_slot, s.subject_name, s.lesson_type, s.teacher, s.room, s.group_name, s.faculty '
                'FROM schedules s '
                'JOIN schedule_weeks sw ON s.id = sw.schedule_id '
                'WHERE s.day_of_week = ? AND sw.week_number = ? '
                'ORDER BY s.time_slot ASC, s.group_name ASC',
                (day_name, week_num),
            )
            rows = cursor.fetchall()

            cursor.execute(
                'SELECT group_name, time_slot FROM cancelled_lessons WHERE target_date = ?',
                (date_str,),
            )
            cancelled_set = {(r[0], r[1]) for r in cursor.fetchall()}

            cursor.execute(
                'SELECT group_name, time_slot, subject_info FROM replacements WHERE target_date = ?',
                (date_str,),
            )
            repl_rows = cursor.fetchall()
            replaced_set = {(r[0], r[1]) for r in repl_rows}

            day_matches = []
            for time_slot, subj, l_type, teacher, room, grp, fac in rows:
                if (grp, time_slot) in cancelled_set or (grp, time_slot) in replaced_set:
                    continue
                if week_num in PRACTICE_WEEKS.get(grp, set()):
                    continue
                teacher_norm = extract_teacher_surname_initials(teacher)
                t_norm_low = (teacher_norm or '').lower()
                t_sur_low = _teacher_surname_only(teacher_norm)
                q_compact = re.sub(r'[.\s]+', '', q_lower)
                t_compact = re.sub(r'[.\s]+', '', t_norm_low)
                if teacher_norm and (
                    q_lower in t_norm_low
                    or (len(q_compact) >= 2 and q_compact in t_compact)
                    or (q_lower == t_sur_low)
                ):
                    fac_tag = get_faculty_short(fac or infer_faculty_by_group(grp))
                    day_matches.append((time_slot, teacher_norm, room, subj, l_type, f'{grp} [{fac_tag}]', False))

            q_surname = q_lower.split()[0] if q_lower.split() else q_lower
            for grp, time_slot, subj_info in repl_rows:
                subj_low = subj_info.lower()
                if q_lower in subj_low or (len(q_surname) >= 3 and q_surname in subj_low):
                    fac_tag = get_faculty_short(infer_faculty_by_group(grp))
                    day_matches.append((time_slot, 'Замена', '', subj_info, 'Замена', f'{grp} [{fac_tag}]', True))

            if day_matches:
                all_day_results.append((dt_obj, day_name, date_str, week_num, day_matches))

    markup = InlineKeyboardMarkup(row_width=2)
    markup.row(
        InlineKeyboardButton('📍 Сегодня', callback_data='tsearch_re_today'),
        InlineKeyboardButton('⏭ Завтра', callback_data='tsearch_re_tomorrow'),
        InlineKeyboardButton('🗓 Неделя', callback_data='tsearch_re_week'),
    )
    markup.row(InlineKeyboardButton('🔍 Новый поиск', callback_data='search_back'))

    set_user_state(chat_id, '', {'last_teacher_query': query})

    if not all_day_results:
        bot.send_message(
            chat_id,
            f'🔎 По запросу <b>«{esc(query)}»</b> за период <b>{period_title}</b> пар не найдено.\n\n'
            f'<i>Попробуй переключить период кнопками ниже:</i>',
            reply_markup=markup,
            parse_mode='HTML',
        )
        return

    text = f'👨‍🏫 <b>ПОИСК: «{esc(query)}»</b>\n📅 <i>{period_title}</i>\n\n'
    for dt_obj, day_name, date_str, week_num, matches in all_day_results:
        if len(target_dates) > 1:
            text += f'🔹 <b><u>{day_name.upper()} ({dt_obj.strftime("%d.%m")})</u></b>\n'
        grouped = {}
        for time_slot, teacher, room, subj, l_type, grp, is_repl in matches:
            key = (time_slot, teacher, room, subj, l_type, is_repl)
            grouped.setdefault(key, []).append(grp)

        for (time_slot, teacher, room, subj, l_type, is_repl), grps in sorted(grouped.items(), key=lambda x: x[0][0]):
            pair_prefix = format_pair_prefix(time_slot)
            grps_str = ', '.join(grps)
            if is_repl:
                text += (
                    f'🔄 {pair_prefix}\n'
                    f'👉 <b>{esc(subj)}</b>\n'
                    f'🎓 <i>Группы: {esc(grps_str)}</i>\n\n'
                )
            else:
                text += (
                    f'{pair_prefix}\n'
                    f'👤 <b>{esc(teacher)}</b>\n'
                    f'📍 <u>{esc(room)}</u>  •  📘 {esc(subj)} <i>({esc(l_type)})</i>\n'
                    f'🎓 <i>Группы: {esc(grps_str)}</i>\n\n'
                )

    chunks = split_long_html_message(text.strip(), max_len=3900)
    for idx, ch in enumerate(chunks):
        kb = markup if idx == len(chunks) - 1 else None
        bot.send_message(chat_id, ch, reply_markup=kb, parse_mode='HTML')


def process_teacher_search(message, payload):
    period = (payload or {}).get('teacher_period', 'today')
    clear_user_state(message.chat.id, keep_payload=True)
    if not message.text or message.text.strip().lower() in ['отмена', '/cancel', '/start']:
        bot.send_message(message.chat.id, '❌ Поиск отменён.')
        return

    query = message.text.strip()
    if len(query) < 2:
        set_user_state(message.chat.id, 'awaiting_teacher_query', {'teacher_period': period})
        bot.send_message(
            message.chat.id,
            '⚠️ Введи хотя бы 2 буквы фамилии преподавателя (или «Отмена»):',
        )
        return

    all_teachers = get_all_unique_teachers()
    q_low = re.sub(r'\s+', ' ', query.lower()).strip()
    q_no_dots = re.sub(r'[.\s]+', '', q_low)

    def _matches_teacher(t_str):
        t_low = t_str.lower()
        t_no_dots = re.sub(r'[.\s]+', '', t_low)
        t_sur = _teacher_surname_only(t_str)
        return (
            q_low in t_low
            or (len(q_no_dots) >= 2 and q_no_dots in t_no_dots)
            or (q_low == t_sur)
        )

    matched_teachers = [t for t in all_teachers if _matches_teacher(t)]

    # Если запрос совпадает с фамилией нескольких преподавателей (например, "Суюндуков" -> "Суюндуков И.В.", "Суюндуков Я.Т.")
    # или найдено несколько вариантов — всегда показываем выбор конкретных преподавателей
    if len(matched_teachers) > 1:
        suggestions = matched_teachers[:8]
        set_user_state(message.chat.id, '', {
            'teacher_suggestions': suggestions,
            'teacher_period': period,
            'teacher_raw_query': query,
        })
        markup = InlineKeyboardMarkup(row_width=1)
        for idx, t_name in enumerate(suggestions):
            markup.add(
                InlineKeyboardButton(
                    f'👤 {t_name}',
                    callback_data=f'tsearch_pick_{idx}',
                )
            )
        markup.add(
            InlineKeyboardButton(
                f'🔎 Искать по всем совпадениям («{short_text(query, 18)}»)',
                callback_data='tsearch_pick_all',
            )
        )
        bot.send_message(
            message.chat.id,
            f'💡 Найдено несколько преподавателей по запросу <b>«{esc(query)}»</b>.\n'
            f'Выбери нужного из списка или выполни общий поиск:',
            reply_markup=markup,
            parse_mode='HTML',
        )
        return

    final_query = matched_teachers[0] if len(matched_teachers) == 1 else query
    execute_teacher_search_for_period(message.chat.id, final_query, period)


def show_free_rooms_day_choice(chat_id, msg_id=None):
    today = now_local()
    tomorrow = today + timedelta(days=1)
    markup = InlineKeyboardMarkup(row_width=1)
    markup.add(
        InlineKeyboardButton(
            f'📍 Сегодня ({DAYS_RUS[today.weekday()]}, {today.strftime("%d.%m")})',
            callback_data='freerooms_day_0',
        ),
        InlineKeyboardButton(
            f'⏭ Завтра ({DAYS_RUS[tomorrow.weekday()]}, {tomorrow.strftime("%d.%m")})',
            callback_data='freerooms_day_1',
        ),
        InlineKeyboardButton('🔙 Назад', callback_data='search_back'),
    )
    text = (
        '🚪 <b>СВОБОДНЫЕ АУДИТОРИИ</b>\n\n'
        '<b>Шаг 1 из 2.</b> Выбери день недели:'
    )
    if msg_id:
        try:
            bot.edit_message_text(text, chat_id, msg_id, reply_markup=markup, parse_mode='HTML')
            return
        except Exception:
            pass
    bot.send_message(chat_id, text, reply_markup=markup, parse_mode='HTML')


def show_free_rooms_pair_choice(chat_id, msg_id=None, day_offset=0):
    target_dt = now_local() + timedelta(days=day_offset)
    _, day_name = get_current_week_and_day(target_dt)
    date_str = target_dt.strftime('%d.%m.%Y')
    day_label = 'Сегодня' if day_offset == 0 else 'Завтра'

    markup = InlineKeyboardMarkup(row_width=2)
    btns = []
    for p_num in range(1, 7):
        emoji = PAIR_EMOJIS.get(p_num, f'{p_num}.')
        slot = NUM_TO_TIME_SLOT[p_num]
        btns.append(
            InlineKeyboardButton(
                f'{emoji} {p_num} пара ({slot[:5]})',
                callback_data=f'freerooms_pair_{day_offset}_{p_num}',
            )
        )
    markup.add(*btns)
    markup.add(InlineKeyboardButton('🔙 Сменить день', callback_data='search_rooms'))

    text = (
        f'🚪 <b>СВОБОДНЫЕ АУДИТОРИИ</b>\n'
        f'📅 <i>{day_label}: {day_name}, {date_str}</i>\n\n'
        f'<b>Шаг 2 из 2.</b> Выбери номер пары:'
    )
    if msg_id:
        try:
            bot.edit_message_text(text, chat_id, msg_id, reply_markup=markup, parse_mode='HTML')
            return
        except Exception:
            pass
    bot.send_message(chat_id, text, reply_markup=markup, parse_mode='HTML')


def extract_rooms_from_text(room_str):
    raw = str(room_str or '').strip()
    if not raw:
        return []
    if raw.lower().startswith('ауд.'):
        rest = raw[4:].strip()
        parts = [p.strip() for p in rest.split(',') if p.strip()]
        return [f'ауд. {p}' for p in parts]
    return [raw]


def show_free_rooms_result(chat_id, msg_id, pair_num, day_offset=0):
    target_dt = now_local() + timedelta(days=day_offset)
    week_num, day_name = get_current_week_and_day(target_dt)
    date_str = target_dt.strftime('%d.%m.%Y')
    time_slot = NUM_TO_TIME_SLOT.get(pair_num, '')

    with db_connection() as conn:
        cursor = conn.cursor()

        cursor.execute('SELECT DISTINCT room FROM schedules')
        all_raw_rooms = [r[0] for r in cursor.fetchall()]

        all_rooms = set()
        for r_str in all_raw_rooms:
            for single in extract_rooms_from_text(r_str):
                if 'конф-зал' not in single.lower() and 'не указана' not in single.lower():
                    all_rooms.add(single)

        cursor.execute(
            'SELECT s.room, s.group_name FROM schedules s '
            'JOIN schedule_weeks sw ON s.id = sw.schedule_id '
            'WHERE s.day_of_week = ? AND sw.week_number = ? AND s.time_slot = ?',
            (day_name, week_num, time_slot),
        )
        busy_rows = cursor.fetchall()

        cursor.execute(
            'SELECT group_name FROM cancelled_lessons WHERE target_date = ? AND time_slot = ?',
            (date_str, time_slot),
        )
        cancelled_groups = {r[0] for r in cursor.fetchall()}

        cursor.execute(
            'SELECT group_name, subject_info FROM replacements WHERE target_date = ? AND time_slot = ?',
            (date_str, time_slot),
        )
        repl_rows_rooms = cursor.fetchall()
        replaced_groups = {r[0] for r in repl_rows_rooms}
        repl_texts = [r[1] for r in repl_rows_rooms]

    busy_rooms = set()
    for r_str, grp in busy_rows:
        if grp in cancelled_groups or grp in replaced_groups:
            continue
        if week_num in PRACTICE_WEEKS.get(grp, set()):
            continue
        for single in extract_rooms_from_text(r_str):
            busy_rooms.add(single)

    for r_txt in repl_texts:
        for room in all_rooms:
            num_part = room.replace('ауд.', '').strip()
            if re.search(rf'(?<!\d){re.escape(num_part)}(?!\d)', r_txt):
                busy_rooms.add(room)

    free_rooms = sorted(all_rooms - busy_rooms)
    busy_sorted = sorted(busy_rooms)

    pair_prefix = format_pair_prefix(time_slot)
    text = (
        f'🚪 <b>СВОБОДНЫЕ АУДИТОРИИ</b>\n'
        f'📅 <i>{day_name}, {date_str} ({week_num}-я нед.)</i>\n'
        f'{pair_prefix}\n\n'
    )

    if day_name == 'Воскресенье':
        text += '🎉 <i>Воскресенье — все аудитории закрыты на выходной.</i>'
    else:
        if free_rooms:
            text += '✅ <b>Свободны по расписанию:</b>\n' + ', '.join(f'<code>{esc(r)}</code>' for r in free_rooms) + '\n\n'
        else:
            text += '⚠️ <i>Свободных аудиторий из списка на эту пару не найдено.</i>\n\n'

        if busy_sorted:
            text += '🔴 <b>Заняты парами:</b>\n' + ', '.join(esc(r) for r in busy_sorted)

    markup = InlineKeyboardMarkup(row_width=1)
    markup.add(
        InlineKeyboardButton('🔄 Выбрать другую пару', callback_data=f'freerooms_day_{day_offset}'),
        InlineKeyboardButton('🔙 В меню поиска', callback_data='search_back'),
    )
    try:
        bot.edit_message_text(text, chat_id, msg_id, reply_markup=markup, parse_mode='HTML')
    except Exception:
        bot.send_message(chat_id, text, reply_markup=markup, parse_mode='HTML')


# --- КОМАНДЫ АДМИНИСТРАТОРА И СТАРОСТ ---
def resolve_target_user(message):
    parts = message.text.split()
    if message.reply_to_message:
        t_id = message.reply_to_message.from_user.id
        t_disp = (
            f'@{message.reply_to_message.from_user.username}'
            if message.reply_to_message.from_user.username
            else str(t_id)
        )
        return t_id, t_disp, None
    if len(parts) > 1:
        arg = parts[1].strip()
        if arg.startswith('@'):
            uname_search = arg[1:].lower()
            with db_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    'SELECT user_id FROM users WHERE LOWER(username) = ?',
                    (uname_search,),
                )
                row = cursor.fetchone()
            if row:
                return row[0], arg, None
            return None, arg, 'not_found'
        if arg.isdigit():
            return int(arg), f'ID {arg}', None
    return None, '', 'no_args'


@bot.message_handler(commands=['starosta', 'addstarosta'])
def cmd_add_starosta(message):
    sync_username(message.from_user)
    if not is_admin_user(message.from_user):
        bot.send_message(
            message.chat.id,
            f'⛔ <b>Доступ запрещён!</b> Выдавать роль старосты может только главный администратор (@{ADMIN_USERNAME}).',
            parse_mode='HTML',
        )
        return

    target_user_id, target_display, err = resolve_target_user(message)
    if err == 'not_found':
        bot.send_message(
            message.chat.id,
            f'⚠️ Пользователь <b>{esc(target_display)}</b> ещё не запускал бота или не найден в базе.\n'
            f'Попроси его нажать /start в боте либо укажи его числовой Telegram ID: <code>/starosta ID</code>',
            parse_mode='HTML',
        )
        return
    if not target_user_id:
        bot.send_message(
            message.chat.id,
            'ℹ️ <b>Инструкция по выдаче прав старосты:</b>\n\n'
            'Используй команду одним из способов:\n'
            '• <code>/starosta @username</code> (если человек уже запускал бота)\n'
            '• <code>/starosta 123456789</code> (по Telegram ID)\n'
            '• Или ответь командой <code>/starosta</code> на сообщение нужного пользователя.',
            parse_mode='HTML',
        )
        return

    with db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            'SELECT course, group_name, subscribed, is_hw_manager, faculty FROM users WHERE user_id = ?',
            (target_user_id,),
        )
        row = cursor.fetchone()
        if row:
            assigned_group = row[1]
            cursor.execute(
                'UPDATE users SET is_starosta = 1, managed_group = ? WHERE user_id = ?',
                (assigned_group, target_user_id),
            )
            sub_state = bool(row[2])
        else:
            default_fac = 'ЕМФ'
            assigned_group = ''
            cursor.execute(
                'INSERT INTO users (user_id, faculty, course, group_name, subscribed, is_starosta, is_hw_manager, notify_updates, managed_group) '
                'VALUES (?, ?, 0, "", 0, 1, 0, 0, "")',
                (target_user_id, default_fac),
            )
            sub_state = False
    invalidate_user_cache(target_user_id)

    bot.send_message(
        message.chat.id,
        f'✅ Пользователю <b>{esc(target_display)}</b> выдана привилегия <b>Старосты</b>!\n'
        f'🛡 <b>Закреплённая группа управления:</b> {esc(assigned_group)}',
        parse_mode='HTML',
    )
    try:
        bot.send_message(
            target_user_id,
            f'🌟 <b>Тебе выдана роль Старосты!</b>\n'
            f'🛡 <b>Закреплённая группа:</b> {esc(assigned_group)}\n\n'
            f'В твоём нижнем меню появилась кнопка <b>«🛠 Спецвозможности»</b> для управления заменами, отменами и домашними заданиями группы.',
            reply_markup=get_main_keyboard(sub_state, is_starosta=True),
            parse_mode='HTML',
        )
    except Exception:
        pass


@bot.message_handler(commands=['unstarosta', 'delstarosta'])
def cmd_remove_starosta(message):
    sync_username(message.from_user)
    if not is_admin_user(message.from_user):
        return

    parts = message.text.split()
    if len(parts) < 2:
        bot.send_message(
            message.chat.id,
            'Укажи <code>@username</code> или <code>ID</code> для снятия роли: <code>/unstarosta @username</code>',
            parse_mode='HTML',
        )
        return

    arg = parts[1].strip()
    with db_connection() as conn:
        cursor = conn.cursor()
        if arg.startswith('@'):
            cursor.execute(
                'UPDATE users SET is_starosta = 0, managed_group = "" WHERE LOWER(username) = ?',
                (arg[1:].lower(),),
            )
        elif arg.isdigit():
            cursor.execute(
                'UPDATE users SET is_starosta = 0, managed_group = "" WHERE user_id = ?',
                (int(arg),),
            )
    invalidate_user_cache()
    bot.send_message(
        message.chat.id,
        f'✅ Роль старосты снята с пользователя <b>{esc(arg)}</b>.',
        parse_mode='HTML',
    )


# --- БЕЗОПАСНЫЙ АВТОМАТИЧЕСКИЙ И РУЧНОЙ БЭКАП БАЗЫ ДАННЫХ ЧЕРЕЗ SQLITE BACKUP API ---
def send_db_backup(manual_chat_id=None):
    if not os.path.exists(DB_PATH):
        if manual_chat_id:
            bot.send_message(manual_chat_id, '⚠️ Файл базы данных schedule.db не найден.')
        return

    if manual_chat_id:
        target_ids = [manual_chat_id]
    else:
        with db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                'SELECT user_id FROM users WHERE LOWER(username) = ?',
                (ADMIN_USERNAME.lower(),),
            )
            target_ids = [r[0] for r in cursor.fetchall()]

    if not target_ids:
        return

    stamp = now_local().strftime('%d.%m.%Y %H:%M')
    caption = f'💾 <b>Резервная копия БД schedule.db</b>\n🕒 <i>{stamp} (UTC+5)</i>'

    tmp_backup_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.db') as tmp_f:
            tmp_backup_path = tmp_f.name

        src_conn = sqlite3.connect(DB_PATH, timeout=30.0)
        dst_conn = sqlite3.connect(tmp_backup_path)
        try:
            src_conn.execute('PRAGMA busy_timeout = 30000')
            src_conn.backup(dst_conn)
            dst_conn.commit()
        finally:
            dst_conn.close()
            src_conn.close()

        for uid in target_ids:
            try:
                with open(tmp_backup_path, 'rb') as db_file:
                    bot.send_document(
                        uid,
                        db_file,
                        visible_file_name=f'schedule_backup_{now_local().strftime("%Y%m%d_%H%M")}.db',
                        caption=caption,
                        parse_mode='HTML',
                    )
            except Exception as e:
                print(f'Ошибка отправки бэкапа БД пользователю {uid}: {e}')
    except Exception as e:
        print(f'Ошибка создания консистентного бэкапа SQLite: {e}')
        if manual_chat_id:
            bot.send_message(manual_chat_id, f'❌ Ошибка при создании бэкапа БД: <code>{esc(e)}</code>', parse_mode='HTML')
    finally:
        if tmp_backup_path and os.path.exists(tmp_backup_path):
            try:
                os.remove(tmp_backup_path)
            except Exception:
                pass


@bot.message_handler(commands=['backup'])
def cmd_backup(message):
    sync_username(message.from_user)
    if not is_admin_user(message.from_user):
        return
    send_db_backup(manual_chat_id=message.chat.id)


# --- ПАНЕЛЬ СТАТИСТИКИ ДЛЯ АДМИНА (/stats) ---
def send_stats_panel(chat_id, user_obj):
    if not is_admin_user(user_obj):
        bot.send_message(
            chat_id,
            '⛔ Эта команда доступна только администратору.',
            parse_mode='HTML',
        )
        return

    with db_connection() as conn:
        cursor = conn.cursor()

        cursor.execute('SELECT COUNT(*) FROM users')
        total_users = cursor.fetchone()[0]

        cursor.execute('SELECT COUNT(*) FROM users WHERE subscribed = 1')
        sub_users = cursor.fetchone()[0]

        cursor.execute('SELECT COUNT(*) FROM users WHERE notify_updates = 1')
        notify_users = cursor.fetchone()[0]

        cursor.execute(
            'SELECT faculty, course, group_name, COUNT(*) FROM users '
            'GROUP BY faculty, course, group_name '
            'ORDER BY faculty ASC, course ASC, COUNT(*) DESC'
        )
        groups_stats = cursor.fetchall()

        cursor.execute(
            'SELECT user_id, username, group_name, managed_group FROM users '
            'WHERE is_starosta = 1 ORDER BY course ASC, group_name ASC'
        )
        starostas = cursor.fetchall()

        cursor.execute('SELECT COUNT(*) FROM replacements')
        repl_count = cursor.fetchone()[0]
        cursor.execute('SELECT COUNT(*) FROM cancelled_lessons')
        canc_count = cursor.fetchone()[0]
        cursor.execute('SELECT COUNT(*) FROM homeworks')
        hw_count = cursor.fetchone()[0]
        cursor.execute('SELECT COUNT(*) FROM schedules')
        sched_rows_count = cursor.fetchone()[0]

    sub_pct = round((sub_users / total_users * 100), 1) if total_users else 0
    notify_pct = round((notify_users / total_users * 100), 1) if total_users else 0
    q_size = MAILING_QUEUE.qsize()
    cur_week, _ = get_current_week_and_day()

    text = (
        '📊 <b>СТАТИСТИКА БОТА СИ УУНиТ</b>\n\n'
        f'📅 <b>Начало 1-й недели года:</b> {ACADEMIC_YEAR_START_DATE.strftime("%d.%m.%Y")} (сейчас <b>{cur_week}-я</b> нед.)\n'
        f'📚 <b>Пар в расписании:</b> {sched_rows_count}\n'
        f'👥 <b>Всего пользователей:</b> {total_users}\n'
        f'🗓 <b>Авторассылка (07:00/18:00):</b> {sub_users} ({sub_pct}%)\n'
        f'🔔 <b>Уведомления (пары/ДЗ):</b> {notify_users} ({notify_pct}%)\n'
        f'🔄 <b>Замен в базе:</b> {repl_count}  •  🚫 <b>Отмен:</b> {canc_count}\n'
        f'📝 <b>Записей ДЗ в базе:</b> {hw_count}\n'
        f'📨 <b>В очереди отправки сейчас:</b> {q_size}\n\n'
        '🎓 <b>Распределение по факультетам и группам:</b>\n'
    )

    if groups_stats:
        for fac, crs, grp, cnt in groups_stats:
            fac_lbl = get_faculty_short(fac or infer_faculty_by_group(grp))
            text += f'• [{esc(fac_lbl)}] {esc(grp)}: <b>{cnt}</b> чел.\n'
    else:
        text += '<i>Пока нет зарегистрированных пользователей.</i>\n'

    text += '\n🌟 <b>Назначенные старосты:</b>\n'
    if starostas:
        for uid, uname, grp, m_grp in starostas:
            label = f'@{uname}' if uname else f'ID <code>{uid}</code>'
            target_g = m_grp or grp
            text += f'• {label} — <i>{esc(target_g)}</i>\n'
    else:
        text += '<i>Нет назначенных старост.</i>\n'

    bot.send_message(chat_id, text, parse_mode='HTML')


@bot.message_handler(commands=['stats'])
def cmd_stats(message):
    sync_username(message.from_user)
    send_stats_panel(message.chat.id, message.from_user)


# --- МАССОВАЯ РАССЫЛКА ОБЪЯВЛЕНИЙ (/broadcast) ---
def open_broadcast_menu(chat_id, user_obj, message_id=None):
    if not is_admin_user(user_obj):
        bot.send_message(
            chat_id,
            '⛔ Команда <code>/broadcast</code> доступна только администратору.',
            parse_mode='HTML',
        )
        return

    clear_user_state(chat_id)
    markup = InlineKeyboardMarkup(row_width=2)
    markup.add(InlineKeyboardButton('📢 Всем пользователям бота', callback_data='bc_target_all'))
    fac_btns = [
        InlineKeyboardButton(f'🏛 {get_faculty_short(f)}', callback_data=f'bc_target_fac_{f}')
        for f in FACULTY_ORDER
    ]
    markup.add(*fac_btns)
    all_courses = sorted({
        c for courses in FACULTY_STRUCTURE.values() for c in courses.keys()
    })
    course_btns = [
        InlineKeyboardButton(f'🎓 {c} курс', callback_data=f'bc_target_{c}')
        for c in all_courses
    ]
    markup.add(*course_btns)
    markup.add(InlineKeyboardButton('🔙 Назад', callback_data='spec_back'))

    text = (
        '📣 <b>МАССОВАЯ РАССЫЛКА ОБЪЯВЛЕНИЙ</b>\n\n'
        'Выбери, кому отправить сообщение от деканата / администрации:'
    )
    if message_id:
        try:
            bot.edit_message_text(text, chat_id, message_id, reply_markup=markup, parse_mode='HTML')
            return
        except Exception:
            pass
    bot.send_message(chat_id, text, reply_markup=markup, parse_mode='HTML')


@bot.message_handler(commands=['broadcast'])
def cmd_broadcast(message):
    sync_username(message.from_user)
    open_broadcast_menu(message.chat.id, message.from_user)


def process_broadcast_content(message, payload):
    target = (payload or {}).get('target', 'all')
    clear_user_state(message.chat.id)
    if message.text and message.text.strip().lower() in ['отмена', '/cancel', '/start']:
        bot.send_message(message.chat.id, '❌ Рассылка отменена.')
        return

    photo_id = message.photo[-1].file_id if message.photo else None
    raw_text = message.caption if photo_id else (message.text or '')

    if not photo_id and not raw_text.strip():
        bot.send_message(
            message.chat.id,
            '⚠️ Поддерживаются только текстовые сообщения или фото с подписью. Рассылка отменена.',
        )
        return

    with db_connection() as conn:
        cursor = conn.cursor()
        if target == 'all':
            cursor.execute('SELECT user_id FROM users')
            target_label = 'всем пользователям'
        elif str(target).startswith('fac_'):
            fac_code = str(target).replace('fac_', '')
            cursor.execute('SELECT user_id FROM users WHERE faculty = ?', (fac_code,))
            target_label = f'факультету {get_faculty_short(fac_code)}'
        else:
            cursor.execute('SELECT user_id FROM users WHERE course = ?', (int(target),))
            target_label = f'{target} курсу'
        recipients = [r[0] for r in cursor.fetchall()]

    header = '📢 <b>ВАЖНОЕ ОБЪЯВЛЕНИЕ</b>\n\n'
    formatted_text = header + esc(raw_text) if raw_text else header.strip()

    for u_id in recipients:
        if photo_id:
            enqueue_hw_media(u_id, [photo_id], formatted_text)
        else:
            enqueue_message(u_id, formatted_text, parse_mode='HTML')

    bot.send_message(
        message.chat.id,
        f'🚀 <b>Рассылка поставлена в фоновую очередь!</b>\n\n'
        f'🎯 <b>Аудитория:</b> {target_label}\n'
        f'👥 <b>Получателей в очереди:</b> {len(recipients)}\n'
        f'⚡ <i>Сообщения плавно доставляются в фоне без блокировки бота.</i>',
        parse_mode='HTML',
    )


# --- INLINE-ВЫБОР ДАТЫ И ПАРЫ ДЛЯ ЗАМЕНЫ / ОТМЕНЫ ---
def build_date_choice_markup(action_prefix):
    today = now_local()
    d0 = today.strftime('%d.%m.%Y')
    d1 = (today + timedelta(days=1)).strftime('%d.%m.%Y')
    d2 = (today + timedelta(days=2)).strftime('%d.%m.%Y')

    markup = InlineKeyboardMarkup(row_width=1)
    markup.add(
        InlineKeyboardButton(f'📍 Сегодня ({d0})', callback_data=f'{action_prefix}_date_{d0}'),
        InlineKeyboardButton(f'⏭ Завтра ({d1})', callback_data=f'{action_prefix}_date_{d1}'),
        InlineKeyboardButton(f'⏩ Послезавтра ({d2})', callback_data=f'{action_prefix}_date_{d2}'),
        InlineKeyboardButton('⌨️ Ввести дату вручную', callback_data=f'{action_prefix}_datemanual'),
        InlineKeyboardButton('❌ Отмена', callback_data='cancel_action'),
    )
    return markup


def show_pair_selection_for_action(chat_id, message_id, action, course, group_name, date_str):
    parsed_date = datetime.strptime(date_str, '%d.%m.%Y').replace(tzinfo=LOCAL_TZ)
    week_num, day_name = get_current_week_and_day(parsed_date)
    slot_map, _ = get_day_lessons_for_date(parsed_date, course, group_name)

    markup = InlineKeyboardMarkup(row_width=1)

    if action == 'del':
        if not slot_map:
            text = (
                f'📅 <b>{day_name}, {date_str}</b> ({week_num}-я нед.)\n\n'
                f'✨ На этот день у группы <b>{esc(group_name)}</b> нет активных пар для отмены.'
            )
            markup.add(InlineKeyboardButton('🔙 Выбрать другую дату', callback_data='del_back_date'))
            markup.add(InlineKeyboardButton('❌ Отмена', callback_data='cancel_action'))
        else:
            text = (
                f'🗑 <b>Отмена пары — {day_name}, {date_str}</b>\n'
                f'<blockquote>🎓 {esc(group_name)}</blockquote>\n\n'
                f'<b>Шаг 2 из 2.</b> Нажми на пару, которую нужно отменить:'
            )
            for p_num in sorted(slot_map.keys()):
                subj_preview = short_text(', '.join(slot_map[p_num]), 34)
                emoji = PAIR_EMOJIS.get(p_num, f'{p_num}.')
                markup.add(
                    InlineKeyboardButton(
                        f'{emoji} {p_num} пара — {subj_preview}',
                        callback_data=f'del_pair_{date_str}_{p_num}',
                    )
                )
            markup.add(InlineKeyboardButton('🔙 Выбрать другую дату', callback_data='del_back_date'))
            markup.add(InlineKeyboardButton('❌ Отмена', callback_data='cancel_action'))
    else:
        text = (
            f'🔄 <b>Добавление замены — {day_name}, {date_str}</b>\n'
            f'<blockquote>🎓 {esc(group_name)}</blockquote>\n\n'
            f'<b>Шаг 2 из 3.</b> Выбери номер пары (1–6):'
        )
        for p_num in range(1, 7):
            emoji = PAIR_EMOJIS.get(p_num, f'{p_num}.')
            if p_num in slot_map:
                subj_preview = short_text(', '.join(slot_map[p_num]), 30)
                btn_text = f'{emoji} {p_num} пара: {subj_preview}'
            else:
                btn_text = f'{emoji} {p_num} пара ({NUM_TO_TIME_SLOT[p_num]}) — свободно'
            markup.add(
                InlineKeyboardButton(
                    btn_text,
                    callback_data=f'repl_pair_{date_str}_{p_num}',
                )
            )
        markup.add(InlineKeyboardButton('🔙 Выбрать другую дату', callback_data='repl_back_date'))
        markup.add(InlineKeyboardButton('❌ Отмена', callback_data='cancel_action'))

    if message_id:
        try:
            bot.edit_message_text(text, chat_id, message_id, reply_markup=markup, parse_mode='HTML')
            return
        except Exception:
            pass
    bot.send_message(chat_id, text, reply_markup=markup, parse_mode='HTML')


def open_add_replacement(chat_id, user_obj, message_id=None):
    clear_user_state(chat_id)
    _, group_name = get_managed_context(user_obj)
    if not group_name:
        bot.send_message(chat_id, '⚠️ Сначала выбери свою группу через /start.')
        return
    if not is_user_starosta(user_obj):
        bot.send_message(
            chat_id,
            '⛔ У тебя нет прав старосты для добавления замен.',
            parse_mode='HTML',
        )
        return

    text = (
        f'📝 <b>Добавление замены для группы:</b>\n'
        f'<blockquote>🎓 {esc(group_name)}</blockquote>\n\n'
        f'<b>Шаг 1 из 3.</b> Выбери <b>дату</b> замены:'
    )
    markup = build_date_choice_markup('repl')
    if message_id:
        try:
            bot.edit_message_text(text, chat_id, message_id, reply_markup=markup, parse_mode='HTML')
            return
        except Exception:
            pass
    bot.send_message(chat_id, text, reply_markup=markup, parse_mode='HTML')


def start_add_replacement(message):
    open_add_replacement(message.chat.id, message.from_user)


def open_delete_pair(chat_id, user_obj, message_id=None):
    clear_user_state(chat_id)
    _, group_name = get_managed_context(user_obj)
    if not group_name:
        bot.send_message(chat_id, '⚠️ Сначала выбери свою группу через /start.')
        return
    if not is_user_starosta(user_obj):
        bot.send_message(
            chat_id,
            '⛔ У тебя нет прав старосты для удаления пар.',
            parse_mode='HTML',
        )
        return

    text = (
        f'🗑 <b>Отмена пары для группы:</b>\n'
        f'<blockquote>🎓 {esc(group_name)}</blockquote>\n\n'
        f'<b>Шаг 1 из 2.</b> Выбери <b>дату</b>, на которую нужно убрать пару:'
    )
    markup = build_date_choice_markup('del')
    if message_id:
        try:
            bot.edit_message_text(text, chat_id, message_id, reply_markup=markup, parse_mode='HTML')
            return
        except Exception:
            pass
    bot.send_message(chat_id, text, reply_markup=markup, parse_mode='HTML')


def start_delete_pair(message):
    open_delete_pair(message.chat.id, message.from_user)


def process_manual_date_for_action(message, payload):
    action = (payload or {}).get('action', '')
    course = (payload or {}).get('course')
    group_name = (payload or {}).get('group_name', '')
    clear_user_state(message.chat.id)
    if not message.text or message.text.strip().lower() in ['отмена', '/cancel', '/start']:
        bot.send_message(message.chat.id, '❌ Действие отменено.')
        return

    date_text = message.text.strip()
    try:
        parsed_date = datetime.strptime(date_text, '%d.%m.%Y')
        date_str = parsed_date.strftime('%d.%m.%Y')
    except ValueError:
        set_user_state(message.chat.id, 'awaiting_manual_date', {
            'action': action,
            'course': course,
            'group_name': group_name,
        })
        bot.send_message(
            message.chat.id,
            '⚠️ <b>Неверный формат даты!</b>\n'
            'Введи дату строго в формате <code>ДД.ММ.ГГГГ</code> (например, <code>21.11.2026</code>) или напиши «Отмена»:',
            parse_mode='HTML',
        )
        return

    if action == 'hw_add':
        show_hw_subject_selection(message.chat.id, None, course, group_name, date_str)
    elif action == 'hw_change_date':
        subj_name = (payload or {}).get('hw_subject', '')
        up_dates = (payload or {}).get('hw_upcoming_dates', [])
        ask_hw_content(message.chat.id, group_name, date_str, subj_name, upcoming_dates=up_dates)
    elif action == 'hw_view':
        send_homeworks_for_date_view(message.chat.id, message.from_user, group_name, date_str)
    else:
        show_pair_selection_for_action(message.chat.id, None, action, course, group_name, date_str)


def process_replacement_subject(message, payload):
    group_name = (payload or {}).get('group_name', '')
    date_str = (payload or {}).get('date_str', '')
    time_slot = (payload or {}).get('time_slot', '')
    clear_user_state(message.chat.id)
    if not message.text or message.text.strip().lower() in ['отмена', '/cancel', '/start']:
        bot.send_message(message.chat.id, '❌ Добавление замены отменено.')
        return

    subject_info = message.text.strip()
    try:
        parsed_date = datetime.strptime(date_str, '%d.%m.%Y').replace(tzinfo=LOCAL_TZ)
    except ValueError:
        bot.send_message(message.chat.id, '⚠️ Ошибка формата даты. Попробуй добавить замену заново.')
        return
    week_number, day_name = get_current_week_and_day(parsed_date)
    course = find_course_by_group(group_name)

    with db_connection() as conn:
        cursor = conn.cursor()

        base_lessons = []
        if course is not None:
            cursor.execute(
                'SELECT s.subject_name, s.lesson_type FROM schedules s '
                'JOIN schedule_weeks sw ON s.id = sw.schedule_id '
                'WHERE s.course_number = ? AND s.group_name = ? '
                'AND s.day_of_week = ? AND sw.week_number = ? AND s.time_slot = ?',
                (course, group_name, day_name, week_number, time_slot),
            )
            base_lessons = [f'{r[0]} ({r[1]})' for r in cursor.fetchall()]

        cursor.execute(
            'DELETE FROM cancelled_lessons WHERE group_name = ? AND target_date = ? AND time_slot = ?',
            (group_name, date_str, time_slot),
        )
        cursor.execute(
            'DELETE FROM replacements WHERE group_name = ? AND target_date = ? AND time_slot = ?',
            (group_name, date_str, time_slot),
        )
        cursor.execute(
            'INSERT INTO replacements (group_name, target_date, time_slot, subject_info, added_by) '
            'VALUES (?, ?, ?, ?, ?)',
            (group_name, date_str, time_slot, subject_info, message.from_user.id),
        )
        cursor.execute(
            'SELECT user_id FROM users WHERE group_name = ? AND notify_updates = 1',
            (group_name,),
        )
        group_users = cursor.fetchall()

    replaced_line = (
        f'<s>{esc(", ".join(base_lessons))}</s> → '
        if base_lessons
        else ''
    )

    pair_prefix = format_pair_prefix(time_slot)
    bot.send_message(
        message.chat.id,
        f'🎉 <b>Замена успешно добавлена и сохранена в расписании!</b>\n\n'
        f'📅 <b>Дата:</b> {date_str}\n'
        f'🎓 <b>Группа:</b> {esc(group_name)}\n'
        f'{pair_prefix}\n'
        f'📘 <b>Предмет:</b> {replaced_line}<b>{esc(subject_info)}</b>',
        parse_mode='HTML',
    )

    notify_text = (
        f'⚡ <b>НОВАЯ ЗАМЕНА В РАСПИСАНИИ!</b>\n'
        f'<blockquote>🎓 <b>Группа:</b> {esc(group_name)}\n'
        f'📅 <b>Дата:</b> {date_str}</blockquote>\n\n'
        f'🔄 {pair_prefix}\n'
        f'👉 {replaced_line}<b>{esc(subject_info)}</b>'
    )
    for (u_id,) in group_users:
        if u_id != message.from_user.id:
            enqueue_message(u_id, notify_text, parse_mode='HTML')


def execute_delete_pair(chat_id, user_id, course, group_name, date_str, time_slot):
    try:
        parsed_date = datetime.strptime(date_str, '%d.%m.%Y').replace(tzinfo=LOCAL_TZ)
    except ValueError:
        bot.send_message(chat_id, '⚠️ Ошибка формата даты. Попробуй отменить пару заново.')
        return
    week_number, day_name = get_current_week_and_day(parsed_date)

    with db_connection() as conn:
        cursor = conn.cursor()

        cursor.execute(
            'SELECT s.subject_name, s.lesson_type FROM schedules s '
            'JOIN schedule_weeks sw ON s.id = sw.schedule_id '
            'WHERE s.course_number = ? AND s.group_name = ? '
            'AND s.day_of_week = ? AND sw.week_number = ? AND s.time_slot = ?',
            (course, group_name, day_name, week_number, time_slot),
        )
        sched_rows = cursor.fetchall()

        cursor.execute(
            'SELECT subject_info FROM replacements '
            'WHERE group_name = ? AND target_date = ? AND time_slot = ?',
            (group_name, date_str, time_slot),
        )
        repl_rows = cursor.fetchall()

        cursor.execute(
            'DELETE FROM replacements WHERE group_name = ? AND target_date = ? AND time_slot = ?',
            (group_name, date_str, time_slot),
        )

        if sched_rows:
            cursor.execute(
                'SELECT id FROM cancelled_lessons WHERE group_name = ? AND target_date = ? AND time_slot = ?',
                (group_name, date_str, time_slot),
            )
            if not cursor.fetchone():
                cursor.execute(
                    'INSERT INTO cancelled_lessons (group_name, target_date, time_slot, cancelled_by) '
                    'VALUES (?, ?, ?, ?)',
                    (group_name, date_str, time_slot, user_id),
                )

        cursor.execute(
            'SELECT user_id FROM users WHERE group_name = ? AND notify_updates = 1',
            (group_name,),
        )
        group_users = cursor.fetchall()

    removed_names = [f'{r[0]} ({r[1]})' for r in sched_rows] + [r[0] for r in repl_rows]
    removed_desc = ', '.join(removed_names) if removed_names else 'Пара по расписанию'

    pair_prefix = format_pair_prefix(time_slot)
    bot.send_message(
        chat_id,
        f'🗑 <b>Пара успешно отменена!</b>\n\n'
        f'📅 <b>Дата:</b> {date_str}\n'
        f'🎓 <b>Группа:</b> {esc(group_name)}\n'
        f'{pair_prefix}\n'
        f'❌ <b>Убрано:</b> {esc(removed_desc)}',
        parse_mode='HTML',
    )

    notify_text = (
        f'🚫 <b>ОТМЕНА ПАРЫ В РАСПИСАНИИ!</b>\n'
        f'<blockquote>🎓 <b>Группа:</b> {esc(group_name)}\n'
        f'📅 <b>Дата:</b> {date_str}</blockquote>\n\n'
        f'❌ {pair_prefix}\n'
        f'<s>{esc(removed_desc)}</s> — <b>ПАРА ОТМЕНЕНА</b>'
    )
    for (u_id,) in group_users:
        if u_id != user_id:
            enqueue_message(u_id, notify_text, parse_mode='HTML')


# --- ПРОСМОТР И ОТКАТ АКТИВНЫХ ЗАМЕН/ОТМЕН ---
def build_active_changes_view(group_name):
    today = now_local().date()
    valid_dates = {(today + timedelta(days=i)).strftime('%d.%m.%Y') for i in range(8)}

    with db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            'SELECT id, target_date, time_slot, subject_info FROM replacements '
            'WHERE group_name = ?',
            (group_name,),
        )
        all_repls = cursor.fetchall()

        cursor.execute(
            'SELECT id, target_date, time_slot FROM cancelled_lessons '
            'WHERE group_name = ?',
            (group_name,),
        )
        all_cancs = cursor.fetchall()

    repls = [r for r in all_repls if r[1] in valid_dates]
    cancs = [c for c in all_cancs if c[1] in valid_dates]

    repls.sort(key=lambda x: (datetime.strptime(x[1], '%d.%m.%Y'), x[2]))
    cancs.sort(key=lambda x: (datetime.strptime(x[1], '%d.%m.%Y'), x[2]))

    text = (
        f'📋 <b>АКТИВНЫЕ ЗАМЕНЫ И ОТМЕНЫ (на 7 дней вперёд)</b>\n'
        f'<blockquote>🎓 <b>Группа:</b> {esc(group_name)}</blockquote>\n\n'
    )
    markup = InlineKeyboardMarkup(row_width=1)

    if not repls and not cancs:
        text += '✨ <i>На ближайшую неделю замен и отмен не внесено.</i>'
        return text, markup

    if repls:
        text += '🔄 <b>Замены:</b>\n'
        for r_id, r_date, r_time, r_subj in repls:
            p_num = get_pair_num(r_time)
            p_label = f'{p_num} пара' if p_num else r_time
            text += f'• <b>{r_date}</b> ({p_label}) — {esc(r_subj)}\n'
            markup.add(
                InlineKeyboardButton(
                    f'❌ Удалить замену: {r_date[:5]} ({p_label}) {short_text(r_subj, 18)}',
                    callback_data=f'rev_repl_{r_id}',
                )
            )
        text += '\n'

    if cancs:
        text += '🚫 <b>Отменённые пары:</b>\n'
        for c_id, c_date, c_time in cancs:
            p_num = get_pair_num(c_time)
            p_label = f'{p_num} пара' if p_num else c_time
            text += f'• <b>{c_date}</b> — {p_label} ({c_time})\n'
            markup.add(
                InlineKeyboardButton(
                    f'↩️ Вернуть пару: {c_date[:5]} ({p_label})',
                    callback_data=f'rev_canc_{c_id}',
                )
            )

    text += '\n<i>Нажми на кнопку ниже, чтобы откатить ошибочную правку в один клик:</i>'
    return text, markup


def show_active_changes(message):
    _, group_name = get_managed_context(message.from_user)
    if not group_name:
        start_onboarding(message)
        return
    if not is_user_starosta(message.from_user):
        bot.send_message(
            message.chat.id,
            '⛔ Просмотр и откат правок доступен только старосте.',
            parse_mode='HTML',
        )
        return

    text, markup = build_active_changes_view(group_name)
    bot.send_message(message.chat.id, text, reply_markup=markup, parse_mode='HTML')


# --- ДОМАШНИЕ ЗАДАНИЯ И РАСПРЕДЕЛЕНИЕ ВОПРОСОВ ---
SHORT_DAYS_RUS = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс']


def find_upcoming_dates_for_subject(course, group_name, subject_query, days_ahead=35, max_dates=4):
    """Находит ближайшие будущие даты (начиная с завтрашнего дня), когда у группы стоит указанный предмет."""
    q_low = str(subject_query or '').strip().lower()
    if not q_low:
        return []
    today = now_local()
    found = []
    for offset in range(1, days_ahead + 1):
        dt = today + timedelta(days=offset)
        _, subjects = get_day_lessons_for_date(dt, course, group_name)
        matched = any(
            q_low == s.strip().lower() or q_low in s.strip().lower() or s.strip().lower() in q_low
            for s in subjects
        )
        if matched:
            found.append(dt)
            if len(found) >= max_dates:
                break
    return found


def open_add_homework(chat_id, user_obj, message_id=None):
    clear_user_state(chat_id)
    course, group_name = get_managed_context(user_obj)
    if not group_name:
        bot.send_message(chat_id, '⚠️ Сначала выбери свою группу через /start.')
        return
    if not can_manage_hw(user_obj):
        bot.send_message(
            chat_id,
            '⚠️ Сначала выбери свой факультет и группу через /start.',
            parse_mode='HTML',
        )
        return

    subjects = get_all_group_subjects(course, group_name)
    subj_entries = []
    for subj in subjects:
        up_dates = find_upcoming_dates_for_subject(course, group_name, subj, days_ahead=35, max_dates=4)
        next_date_str = up_dates[0].strftime('%d.%m.%Y') if up_dates else ''
        next_badge = (
            f' ({SHORT_DAYS_RUS[up_dates[0].weekday()]} {up_dates[0].strftime("%d.%m")})'
            if up_dates
            else ''
        )
        subj_entries.append({
            'subject': subj,
            'next_date': next_date_str,
            'upcoming_dates': [d.strftime('%d.%m.%Y') for d in up_dates],
            'badge': next_badge,
        })

    set_user_state(chat_id, 'awaiting_hw_manual_subj', {
        'hw_group': group_name,
        'hw_course': course,
        'hw_subj_entries': subj_entries,
    })

    markup = InlineKeyboardMarkup(row_width=1)
    for idx, ent in enumerate(subj_entries):
        badge = ent['badge']
        max_s_len = max(16, 38 - len(badge))
        btn_label = f'📘 {short_text(ent["subject"], max_s_len)}{badge}'
        markup.add(
            InlineKeyboardButton(btn_label, callback_data=f'hwadd_picksubj_{idx}')
        )
    markup.add(InlineKeyboardButton('📅 Выбрать сначала дату', callback_data='hwadd_bydate_menu'))
    markup.add(InlineKeyboardButton('❌ Отмена', callback_data='cancel_action'))

    text = (
        f'📚 <b>ДОБАВЛЕНИЕ ДОМАШНЕГО ЗАДАНИЯ</b>\n'
        f'<blockquote>🎓 <b>Группа:</b> {esc(group_name)}</blockquote>\n\n'
        f'<b>Шаг 1 из 3.</b> Нажми на предмет из списка ниже <b>или просто напиши его название сообщением</b> — '
        f'бот сам прикрепит ДЗ к ближайшей следующей паре по расписанию:'
    )
    if message_id:
        try:
            bot.edit_message_text(text, chat_id, message_id, reply_markup=markup, parse_mode='HTML')
            return
        except Exception:
            pass
    bot.send_message(chat_id, text, reply_markup=markup, parse_mode='HTML')


def open_add_homework_by_date(chat_id, user_obj, message_id=None):
    clear_user_state(chat_id)
    _, group_name = get_managed_context(user_obj)
    if not group_name:
        bot.send_message(chat_id, '⚠️ Сначала выбери свою группу через /start.')
        return

    today = now_local()
    d1 = (today + timedelta(days=1)).strftime('%d.%m.%Y')
    d2 = (today + timedelta(days=2)).strftime('%d.%m.%Y')
    d3 = (today + timedelta(days=3)).strftime('%d.%m.%Y')
    d7 = (today + timedelta(days=7)).strftime('%d.%m.%Y')

    markup = InlineKeyboardMarkup(row_width=1)
    markup.add(
        InlineKeyboardButton(f'⏭ На завтра ({d1})', callback_data=f'hwadd_date_{d1}'),
        InlineKeyboardButton(f'⏩ На послезавтра ({d2})', callback_data=f'hwadd_date_{d2}'),
        InlineKeyboardButton(f'📅 Через 3 дня ({d3})', callback_data=f'hwadd_date_{d3}'),
        InlineKeyboardButton(f'🗓 Через неделю ({d7})', callback_data=f'hwadd_date_{d7}'),
        InlineKeyboardButton('⌨️ Ввести дату вручную', callback_data='hwadd_datemanual'),
        InlineKeyboardButton('🔙 К выбору по предмету', callback_data='hw_open_add'),
    )

    text = (
        f'📚 <b>ДОБАВЛЕНИЕ ДЗ ПО ДАТЕ</b>\n'
        f'<blockquote>🎓 <b>Группа:</b> {esc(group_name)}</blockquote>\n\n'
        f'Выбери дату сдачи (дедлайн):'
    )
    if message_id:
        try:
            bot.edit_message_text(text, chat_id, message_id, reply_markup=markup, parse_mode='HTML')
            return
        except Exception:
            pass
    bot.send_message(chat_id, text, reply_markup=markup, parse_mode='HTML')


def start_add_homework(message):
    open_add_homework(message.chat.id, message.from_user)


def show_hw_subject_selection(chat_id, message_id, course, group_name, date_str, show_all=False):
    parsed_date = datetime.strptime(date_str, '%d.%m.%Y').replace(tzinfo=LOCAL_TZ)
    _, day_name = get_current_week_and_day(parsed_date)

    if show_all:
        subjects = get_all_group_subjects(course, group_name)
        header_note = 'Все предметы группы'
    else:
        _, subjects = get_day_lessons_for_date(parsed_date, course, group_name)
        header_note = f'Пары на {day_name}, {date_str}'

    set_user_state(chat_id, '', {
        'hw_date': date_str,
        'hw_subjects': subjects,
        'hw_group': group_name,
        'hw_course': course,
    })

    markup = InlineKeyboardMarkup(row_width=1)
    for idx, subj in enumerate(subjects):
        markup.add(
            InlineKeyboardButton(
                f'📘 {short_text(subj, 38)}',
                callback_data=f'hwadd_subj_{idx}',
            )
        )

    if not show_all:
        markup.add(InlineKeyboardButton('📖 Все предметы группы', callback_data='hwadd_allsubj'))
    markup.add(InlineKeyboardButton('✍️ Ввести предмет вручную', callback_data='hwadd_subjmanual'))
    markup.add(InlineKeyboardButton('❌ Отмена', callback_data='cancel_action'))

    text = (
        f'📚 <b>ДОБАВЛЕНИЕ ДЗ — {date_str} ({day_name})</b>\n\n'
        f'Выбери предмет <i>({header_note})</i>:'
    )
    if not subjects and not show_all:
        text += '\n\n<i>На этот день пар по расписанию нет — выбери из всех предметов группы или введи вручную.</i>'

    if message_id:
        try:
            bot.edit_message_text(text, chat_id, message_id, reply_markup=markup, parse_mode='HTML')
            return
        except Exception:
            pass
    bot.send_message(chat_id, text, reply_markup=markup, parse_mode='HTML')


def format_queue_slots_block(queue_id, viewer_user_id=None, header_title='📋 Распределение вопросов:'):
    if not queue_id:
        return ''
    with db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('SELECT max_slots, is_open FROM queues WHERE id = ?', (queue_id,))
        q_row = cursor.fetchone()
        if not q_row:
            return ''
        max_slots, is_open = q_row
        cursor.execute(
            'SELECT slot_num, user_id, user_label FROM queue_slots WHERE queue_id = ? ORDER BY slot_num ASC',
            (queue_id,),
        )
        slots = cursor.fetchall()

    taken_by_num = {s_num: (u_id, u_lbl) for s_num, u_id, u_lbl in slots}
    lines = [f'\n\n<b>{esc(header_title)}</b> <i>({len(slots)}/{max_slots})</i>']
    for num in range(1, max_slots + 1):
        if num in taken_by_num:
            u_id, u_lbl = taken_by_num[num]
            me_mark = ' 👈 <b>(ты)</b>' if (viewer_user_id and u_id == viewer_user_id) else ''
            lines.append(f'<b>{num}.</b> {esc(u_lbl)}{me_mark}')
        else:
            lines.append(f'<b>{num}.</b> <i>— свободно —</i>')
    return '\n'.join(lines)


def build_queue_slot_buttons(
    queue_id,
    viewer_user_id=None,
    can_moderate=False,
    include_manage_buttons=None,
):
    if not queue_id:
        return include_manage_buttons

    with db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('SELECT max_slots, is_open FROM queues WHERE id = ?', (queue_id,))
        q_row = cursor.fetchone()
        if not q_row:
            return include_manage_buttons
        max_slots, is_open = q_row
        cursor.execute(
            'SELECT slot_num, user_id, user_label FROM queue_slots WHERE queue_id = ? ORDER BY slot_num ASC',
            (queue_id,),
        )
        slots = cursor.fetchall()

    taken_map = {s_num: (u_id, u_lbl) for s_num, u_id, u_lbl in slots}
    my_slots = [s_num for s_num, u_id, _ in slots if viewer_user_id and u_id == viewer_user_id]

    markup = InlineKeyboardMarkup(row_width=5)
    if is_open:
        free_nums = [n for n in range(1, max_slots + 1) if n not in taken_map]
        if free_nums:
            slot_btns = [
                InlineKeyboardButton(f'➕ №{n}', callback_data=f'hwq_pick_{queue_id}_{n}')
                for n in free_nums[:35]
            ]
            for i in range(0, len(slot_btns), 5):
                markup.row(*slot_btns[i : i + 5])

        removable_nums = sorted(taken_map.keys()) if can_moderate else my_slots
        if removable_nums:
            rem_btns = []
            for n in removable_nums[:25]:
                owner_id, owner_lbl = taken_map[n]
                if viewer_user_id and owner_id == viewer_user_id:
                    btn_txt = f'❌ Снять №{n}'
                else:
                    short_owner = short_text(owner_lbl.split('(')[0].strip(), 10)
                    btn_txt = f'🧹 №{n} ({short_owner})'
                rem_btns.append(
                    InlineKeyboardButton(btn_txt, callback_data=f'hwq_free_{queue_id}_{n}')
                )
            for i in range(0, len(rem_btns), 3):
                markup.row(*rem_btns[i : i + 3])

        markup.row(
            InlineKeyboardButton('🔄 Обновить список', callback_data=f'hwq_refresh_{queue_id}')
        )

    if include_manage_buttons:
        for row_btns in include_manage_buttons:
            markup.row(*row_btns)
    return markup


def ask_hw_content(chat_id, group_name, date_str, subject_name, upcoming_dates=None):
    up_list = list(upcoming_dates or [])
    if date_str and date_str not in up_list:
        up_list.insert(0, date_str)

    set_user_state(chat_id, 'awaiting_hw_content', {
        'hw_group': group_name,
        'hw_date': date_str,
        'hw_subject': subject_name,
        'hw_upcoming_dates': up_list,
    })

    try:
        dt_obj = datetime.strptime(date_str, '%d.%m.%Y').replace(tzinfo=LOCAL_TZ)
        day_label = f' ({DAYS_RUS[dt_obj.weekday()]})'
    except Exception:
        day_label = ''

    markup = None
    alt_dates = [d for d in up_list if d != date_str][:3]
    if alt_dates or True:
        markup = InlineKeyboardMarkup(row_width=2)
        if alt_dates:
            alt_btns = []
            for d_alt in alt_dates:
                try:
                    dt_a = datetime.strptime(d_alt, '%d.%m.%Y')
                    lbl_a = f'📅 {SHORT_DAYS_RUS[dt_a.weekday()]} {d_alt[:5]}'
                except Exception:
                    lbl_a = f'📅 {d_alt[:5]}'
                alt_btns.append(InlineKeyboardButton(lbl_a, callback_data=f'hwadd_setdate_{d_alt}'))
            markup.add(*alt_btns)
        markup.row(
            InlineKeyboardButton('⌨️ Другая дата', callback_data='hwadd_changedate_manual'),
            InlineKeyboardButton('❌ Отмена', callback_data='cancel_action'),
        )

    bot.send_message(
        chat_id,
        f'📘 <b>Предмет:</b> {esc(subject_name)}\n'
        f'📅 <b>Дата пары (дедлайн):</b> <b>{date_str}</b>{day_label}\n\n'
        f'<b>Шаг 2 из 3.</b> Отправь текст домашнего задания (можно сразу с фото, альбомом или документом):\n'
        f'<i>Если нужно прикрепить к следующей паре по этому предмету или другой дате — нажми кнопку ниже.</i>',
        reply_markup=markup,
        parse_mode='HTML',
    )


def process_hw_manual_subject(message, payload):
    group_name = (payload or {}).get('hw_group', '')
    course = (payload or {}).get('hw_course') or find_course_by_group(group_name)
    fixed_date_str = (payload or {}).get('hw_date', '')
    clear_user_state(message.chat.id)
    if not message.text or message.text.strip().lower() in ['отмена', '/cancel', '/start']:
        bot.send_message(message.chat.id, '❌ Добавление ДЗ отменено.')
        return

    raw_subj = message.text.strip()
    all_subjs = get_all_group_subjects(course, group_name) if course else []
    matched_subj = raw_subj
    for s in all_subjs:
        if s.strip().lower() == raw_subj.lower():
            matched_subj = s
            break
    else:
        partial = [s for s in all_subjs if raw_subj.lower() in s.strip().lower()]
        if len(partial) == 1:
            matched_subj = partial[0]

    if fixed_date_str:
        ask_hw_content(message.chat.id, group_name, fixed_date_str, matched_subj)
        return

    up_dts = find_upcoming_dates_for_subject(course, group_name, matched_subj, days_ahead=35, max_dates=4) if course else []
    up_strs = [d.strftime('%d.%m.%Y') for d in up_dts]
    if up_strs:
        ask_hw_content(message.chat.id, group_name, up_strs[0], matched_subj, upcoming_dates=up_strs)
        return

    # Если предмет не найден в расписании на ближайшие 5 недель — предлагаем быстрый выбор даты
    today = now_local()
    d1 = (today + timedelta(days=1)).strftime('%d.%m.%Y')
    d2 = (today + timedelta(days=2)).strftime('%d.%m.%Y')
    d7 = (today + timedelta(days=7)).strftime('%d.%m.%Y')
    set_user_state(message.chat.id, '', {
        'hw_group': group_name,
        'hw_course': course,
        'hw_pending_custom_subj': matched_subj,
    })
    markup = InlineKeyboardMarkup(row_width=1)
    markup.add(
        InlineKeyboardButton(f'⏭ На завтра ({d1})', callback_data=f'hwadd_custdate_{d1}'),
        InlineKeyboardButton(f'⏩ На послезавтра ({d2})', callback_data=f'hwadd_custdate_{d2}'),
        InlineKeyboardButton(f'🗓 Через неделю ({d7})', callback_data=f'hwadd_custdate_{d7}'),
        InlineKeyboardButton('⌨️ Ввести дату вручную', callback_data='hwadd_changedate_manual_new'),
        InlineKeyboardButton('❌ Отмена', callback_data='cancel_action'),
    )
    bot.send_message(
        message.chat.id,
        f'📘 <b>Предмет:</b> {esc(matched_subj)}\n\n'
        f'⚠️ Этот предмет не нашёлся в ближайшем расписании группы. Выбери дату, к которой прикрепить ДЗ:',
        reply_markup=markup,
        parse_mode='HTML',
    )


def ask_hw_attach_queue(chat_id, author_id, author_label, group_name, date_str, subject_name, hw_text, photo_ids, doc_file_id='', doc_file_name=''):
    set_user_state(chat_id, '', {
        'pending_hw_draft': {
            'author_id': author_id,
            'author_label': author_label,
            'group_name': group_name,
            'date_str': date_str,
            'subject_name': subject_name,
            'hw_text': hw_text,
            'photo_ids': photo_ids,
            'doc_file_id': doc_file_id,
            'doc_file_name': doc_file_name,
        }
    })
    markup = InlineKeyboardMarkup(row_width=1)
    markup.add(
        InlineKeyboardButton('✅ Да, добавить слоты для разбора вопросов', callback_data='hw_queue_yes'),
        InlineKeyboardButton('⏭ Нет, опубликовать ДЗ без слотов', callback_data='hw_queue_no'),
        InlineKeyboardButton('❌ Отмена', callback_data='cancel_action'),
    )
    bot.send_message(
        chat_id,
        f'📋 <b>Шаг 3 из 3. Слоты для разбора вопросов</b>\n\n'
        f'Нужно ли прикрепить к этому ДЗ по предмету <b>«{esc(subject_name)}»</b> нумерованные слоты (1..N), чтобы студенты могли разобрать вопросы (в том числе по несколько вопросов на человека)?',
        reply_markup=markup,
        parse_mode='HTML',
    )


def process_hw_queue_count_input(message, payload):
    clear_user_state(message.chat.id, keep_payload=True)

    if not message.text or message.text.strip().lower() in ['отмена', '/cancel', '/start']:
        clear_user_state(message.chat.id)
        bot.send_message(message.chat.id, '❌ Добавление ДЗ отменено.')
        return

    raw = message.text.strip()
    if not raw.isdigit() or not (1 <= int(raw) <= 60):
        set_user_state(message.chat.id, 'awaiting_hw_slots_count', payload)
        bot.send_message(
            message.chat.id,
            '⚠️ Введи целое число от <b>1</b> до <b>60</b> (например, <code>12</code> или <code>20</code>), либо напиши «Отмена»:',
            parse_mode='HTML',
        )
        return

    slots_count = int(raw)
    draft = (payload or {}).get('pending_hw_draft')
    if not draft:
        clear_user_state(message.chat.id)
        bot.send_message(message.chat.id, '⚠️ Сессия создания ДЗ устарела. Начни заново.')
        return

    save_and_broadcast_homework(
        message.chat.id,
        draft['author_id'],
        draft['group_name'],
        draft['date_str'],
        draft['subject_name'],
        draft['hw_text'],
        draft['photo_ids'],
        doc_file_id=draft['doc_file_id'],
        doc_file_name=draft['doc_file_name'],
        queue_slots_count=slots_count,
        author_label=draft.get('author_label') or format_hw_author_label(message.from_user),
    )


def save_and_broadcast_homework(
    chat_id,
    author_id,
    group_name,
    date_str,
    subject_name,
    hw_text,
    photo_ids,
    doc_file_id='',
    doc_file_name='',
    queue_slots_count=None,
    author_label='',
):
    clear_user_state(chat_id)
    serialized_photos = serialize_photo_ids(photo_ids)
    created_stamp = now_local().strftime('%d.%m.%Y %H:%M')

    if not author_label:
        with db_connection() as conn:
            c_tmp = conn.cursor()
            c_tmp.execute('SELECT username FROM users WHERE user_id = ?', (author_id,))
            u_r = c_tmp.fetchone()
            author_label = f'@{u_r[0]}' if (u_r and u_r[0]) else f'ID {author_id}'

    with db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            'INSERT INTO homeworks '
            '(group_name, target_date, subject_name, hw_text, photo_file_id, doc_file_id, doc_file_name, added_by, created_at) '
            'VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)',
            (
                group_name,
                date_str,
                subject_name,
                hw_text,
                serialized_photos,
                doc_file_id,
                doc_file_name,
                author_id,
                created_stamp,
            ),
        )
        hw_id = cursor.lastrowid
        queue_id = None

        if queue_slots_count and queue_slots_count > 0:
            q_title = f'{subject_name} (ДЗ на {date_str})'
            cursor.execute(
                'INSERT INTO queues (group_name, title, max_slots, is_open, homework_id, created_by, created_at) '
                'VALUES (?, ?, ?, 1, ?, ?, ?)',
                (group_name, q_title, queue_slots_count, hw_id, author_id, created_stamp),
            )
            queue_id = cursor.lastrowid
            cursor.execute(
                'UPDATE homeworks SET queue_id = ? WHERE id = ?',
                (queue_id, hw_id),
            )

        cursor.execute(
            'SELECT user_id, is_starosta, is_hw_manager, username FROM users WHERE group_name = ? AND notify_updates = 1',
            (group_name,),
        )
        group_users = cursor.fetchall()

    att_notes = []
    if photo_ids:
        att_notes.append(f'фото: {len(photo_ids)} шт.')
    if doc_file_id:
        att_notes.append(f'файл: {esc(doc_file_name or "документ")}')
    att_fallback = f'<i>(прикреплено: {", ".join(att_notes)})</i>' if att_notes else ''

    q_block_author = format_queue_slots_block(queue_id, viewer_user_id=author_id) if queue_id else ''
    author_markup = (
        build_queue_slot_buttons(queue_id, viewer_user_id=author_id, can_moderate=True)
        if queue_id
        else None
    )

    confirm_text = (
        f'✅ <b>Домашнее задание сохранено!</b>\n\n'
        f'📅 <b>Дедлайн:</b> {date_str}\n'
        f'📘 <b>Предмет:</b> {esc(subject_name)}\n'
        f'📝 <b>Задание:</b> {esc(hw_text) if hw_text else att_fallback}\n'
        f'👤 <b>Добавил:</b> {esc(author_label)}'
        f'{q_block_author}'
    )
    send_hw_media(
        chat_id,
        photo_ids,
        confirm_text,
        reply_markup=author_markup,
        doc_file_id=doc_file_id,
        doc_file_name=doc_file_name,
    )

    for u_id, u_st, u_hw, u_uname in group_users:
        if u_id != author_id:
            is_u_mod = bool(u_st == 1 or (u_uname or '').lower() == ADMIN_USERNAME.lower())
            q_block_user = format_queue_slots_block(queue_id, viewer_user_id=u_id) if queue_id else ''
            user_markup = (
                build_queue_slot_buttons(queue_id, viewer_user_id=u_id, can_moderate=is_u_mod)
                if queue_id
                else None
            )
            notify_text = (
                f'📚 <b>НОВОЕ ДОМАШНЕЕ ЗАДАНИЕ!</b>\n'
                f'<blockquote>🎓 <b>Группа:</b> {esc(group_name)}\n'
                f'📅 <b>Срок сдачи:</b> {date_str}\n'
                f'👤 <b>Добавил:</b> {esc(author_label)}</blockquote>\n\n'
                f'📘 <b>{esc(subject_name)}</b>\n'
                f'📝 {esc(hw_text) if hw_text else "<i>(см. прикреплённое вложение)</i>"}'
                f'{q_block_user}'
            )
            enqueue_hw_media(
                u_id,
                photo_ids,
                notify_text,
                reply_markup=user_markup,
                doc_file_id=doc_file_id,
                doc_file_name=doc_file_name,
            )


def flush_media_group_hw(mg_id):
    try:
        with media_group_lock:
            entry = media_group_buffer.pop(mg_id, None)
        if not entry:
            return
        sorted_items = sorted(entry['items'], key=lambda x: x[0])
        photo_ids = [x[1] for x in sorted_items]
        captions = [x[2] for x in sorted_items if x[2]]
        hw_text = captions[0] if captions else ''

        ask_hw_attach_queue(
            entry['chat_id'],
            entry['author_id'],
            entry['author_label'],
            entry['group_name'],
            entry['date_str'],
            entry['subject_name'],
            hw_text,
            photo_ids,
        )
    except Exception as e:
        print(f'Ошибка при сборке альбома ДЗ ({mg_id}): {e}')


@bot.message_handler(
    content_types=['photo'],
    func=lambda m: bool(m.media_group_id and str(m.media_group_id) in media_group_buffer),
)
def collect_subsequent_album_photos(message):
    mg_id = str(message.media_group_id)
    with media_group_lock:
        entry = media_group_buffer.get(mg_id)
        if not entry:
            return
        entry['items'].append(
            (message.message_id, message.photo[-1].file_id, (message.caption or '').strip())
        )
        if entry.get('timer'):
            entry['timer'].cancel()
        t = threading.Timer(1.3, flush_media_group_hw, args=[mg_id])
        entry['timer'] = t
        t.start()


def process_hw_content(message, payload):
    group_name = (payload or {}).get('hw_group', '')
    date_str = (payload or {}).get('hw_date', '')
    subject_name = (payload or {}).get('hw_subject', '')
    clear_user_state(message.chat.id)
    if message.text and message.text.strip().lower() in ['отмена', '/cancel', '/start']:
        bot.send_message(message.chat.id, '❌ Добавление ДЗ отменено.')
        return

    author_label = format_hw_author_label(message.from_user)

    if message.photo and message.media_group_id:
        mg_id = str(message.media_group_id)
        with media_group_lock:
            entry = media_group_buffer.get(mg_id)
            if not entry:
                entry = {
                    '_created_ts': time.time(),
                    'chat_id': message.chat.id,
                    'author_id': message.from_user.id,
                    'author_label': author_label,
                    'group_name': group_name,
                    'date_str': date_str,
                    'subject_name': subject_name,
                    'items': [],
                    'timer': None,
                }
                media_group_buffer[mg_id] = entry
            entry['items'].append(
                (message.message_id, message.photo[-1].file_id, (message.caption or '').strip())
            )
            if entry.get('timer'):
                entry['timer'].cancel()
            t = threading.Timer(1.3, flush_media_group_hw, args=[mg_id])
            entry['timer'] = t
            t.start()
        return

    photo_ids = [message.photo[-1].file_id] if message.photo else []
    doc_file_id = message.document.file_id if message.document else ''
    doc_file_name = (message.document.file_name or 'Документ') if message.document else ''
    has_media = bool(photo_ids or doc_file_id)
    hw_text = (message.caption or '').strip() if has_media else (message.text or '').strip()

    if not has_media and not hw_text:
        set_user_state(message.chat.id, 'awaiting_hw_content', {
            'hw_group': group_name,
            'hw_date': date_str,
            'hw_subject': subject_name,
        })
        bot.send_message(
            message.chat.id,
            '⚠️ Пожалуйста, отправь текст задания, фото или документ (можно с подписью). Для отмены напиши «Отмена»:',
        )
        return

    ask_hw_attach_queue(
        message.chat.id,
        message.from_user.id,
        author_label,
        group_name,
        date_str,
        subject_name,
        hw_text,
        photo_ids,
        doc_file_id=doc_file_id,
        doc_file_name=doc_file_name,
    )


def process_hw_edit_text(message, payload):
    hw_id = (payload or {}).get('hw_id')
    group_name = (payload or {}).get('group_name', '')
    clear_user_state(message.chat.id)
    if not message.text or message.text.strip().lower() in ['отмена', '/cancel', '/start']:
        bot.send_message(message.chat.id, '❌ Редактирование ДЗ отменено.')
        return

    new_text = message.text.strip()
    with db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            'SELECT target_date, subject_name FROM homeworks WHERE id = ? AND group_name = ?',
            (hw_id, group_name),
        )
        row = cursor.fetchone()
        if not row:
            bot.send_message(message.chat.id, '⚠️ Запись не найдена (возможно, она уже была удалена).')
            return
        hw_date, hw_subj = row
        cursor.execute(
            'UPDATE homeworks SET hw_text = ? WHERE id = ? AND group_name = ?',
            (new_text, hw_id, group_name),
        )
        cursor.execute(
            'SELECT user_id FROM users WHERE group_name = ? AND notify_updates = 1',
            (group_name,),
        )
        group_users = cursor.fetchall()

    bot.send_message(
        message.chat.id,
        f'✅ <b>Текст ДЗ обновлён!</b>\n\n'
        f'📅 <b>Дедлайн:</b> {hw_date}\n'
        f'📘 <b>Предмет:</b> {esc(hw_subj)}\n'
        f'📝 <b>Новый текст:</b> {esc(new_text)}',
        parse_mode='HTML',
    )

    notify_text = (
        f'✏️ <b>ОБНОВЛЁН ТЕКСТ ДОМАШНЕГО ЗАДАНИЯ</b>\n'
        f'<blockquote>🎓 <b>Группа:</b> {esc(group_name)}\n'
        f'📅 <b>Срок сдачи:</b> {hw_date}</blockquote>\n\n'
        f'📘 <b>{esc(hw_subj)}</b>\n'
        f'📝 {esc(new_text)}'
    )
    for (u_id,) in group_users:
        if u_id != message.from_user.id:
            enqueue_message(u_id, notify_text, parse_mode='HTML')


def open_homework_menu(message):
    user = get_user_data(message.from_user.id)
    if not user:
        start_onboarding(message)
        return

    tomorrow_str = (now_local() + timedelta(days=1)).strftime('%d.%m.%Y')
    markup = InlineKeyboardMarkup(row_width=1)
    markup.add(
        InlineKeyboardButton(f'⏭ ДЗ на завтра ({tomorrow_str})', callback_data=f'hw_showdate_{tomorrow_str}'),
        InlineKeyboardButton('📆 ДЗ на ближайшую неделю', callback_data='hw_view_week'),
        InlineKeyboardButton('🗓 Все актуальные ДЗ', callback_data='hw_view_all'),
        InlineKeyboardButton('🔍 Выбрать конкретную дату', callback_data='hw_view_manual'),
    )
    if can_manage_hw(message.from_user):
        markup.add(InlineKeyboardButton('➕ Добавить новое ДЗ', callback_data='hw_open_add'))

    bot.send_message(
        message.chat.id,
        f'📚 <b>ДОМАШНИЕ ЗАДАНИЯ И ДЕДЛАЙНЫ</b>\n'
        f'<blockquote>🎓 <b>Группа:</b> {esc(user[1])}</blockquote>\n\n'
        f'Выбери период для просмотра:',
        reply_markup=markup,
        parse_mode='HTML',
    )


def format_user_queue_label(user_obj):
    full_name = ' '.join(
        p for p in [user_obj.first_name, user_obj.last_name] if p
    ).strip()
    if user_obj.username:
        return f'{full_name} (@{user_obj.username})' if full_name else f'@{user_obj.username}'
    return full_name or f'Студент ID {user_obj.id}'


def format_hw_author_label(user_obj):
    if user_obj and getattr(user_obj, 'username', None):
        return f'@{user_obj.username}'
    if user_obj:
        full_name = ' '.join(
            p for p in [user_obj.first_name, user_obj.last_name] if p
        ).strip()
        return full_name or f'ID {user_obj.id}'
    return ''


def can_edit_specific_hw(viewer_user_obj, target_group_name, hw_added_by=None):
    if is_admin_user(viewer_user_obj):
        return True
    if hw_added_by and int(hw_added_by) == int(viewer_user_obj.id):
        return True
    if is_user_starosta(viewer_user_obj):
        _, managed_grp = get_managed_context(viewer_user_obj)
        return bool(target_group_name and managed_grp == target_group_name)
    return False


def get_author_display_by_id(author_id):
    if not author_id:
        return ''
    with db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('SELECT username FROM users WHERE user_id = ?', (int(author_id),))
        row = cursor.fetchone()
    if row and row[0]:
        return f'@{row[0]}'
    return f'ID {author_id}'


def render_hw_card_payload(hw_row, viewer_user_obj, target_group_name=None):
    hw_id, hw_date, hw_subj, hw_text, hw_photo, hw_doc_id, hw_doc_name, hw_queue_id = hw_row[:8]
    hw_added_by = hw_row[8] if len(hw_row) > 8 else None
    can_edit = can_edit_specific_hw(viewer_user_obj, target_group_name, hw_added_by)
    q_block = format_queue_slots_block(hw_queue_id, viewer_user_id=viewer_user_obj.id) if hw_queue_id else ''
    author_disp = get_author_display_by_id(hw_added_by)
    author_line = f'\n👤 <i>Добавил: {esc(author_disp)}</i>' if author_disp else ''
    item_text = (
        f'📅 <b>Дедлайн:</b> {hw_date}\n'
        f'📘 <b>{esc(hw_subj)}</b>\n'
        f'📝 {esc(hw_text) if hw_text else "<i>(см. вложение)</i>"}'
        f'{author_line}'
        f'{q_block}'
    )
    manage_rows = []
    if can_edit:
        manage_rows.append([
            InlineKeyboardButton('✏️ Изменить текст', callback_data=f'hw_edit_{hw_id}'),
            InlineKeyboardButton('🗑 Удалить ДЗ', callback_data=f'hw_del_{hw_id}'),
        ])

    if hw_queue_id:
        markup = build_queue_slot_buttons(
            hw_queue_id,
            viewer_user_id=viewer_user_obj.id,
            can_moderate=can_edit,
            include_manage_buttons=manage_rows,
        )
    elif manage_rows:
        markup = InlineKeyboardMarkup(row_width=2)
        for r in manage_rows:
            markup.row(*r)
    else:
        markup = None
    return item_text, markup


def send_hw_items(chat_id, user_obj, items, title, target_group_name=None):
    if not items:
        bot.send_message(
            chat_id,
            f'📚 <b>{title}</b>\n\n✨ <i>Заданий на выбранный период пока нет!</i>',
            parse_mode='HTML',
        )
        return

    bot.send_message(chat_id, f'📚 <b>{title}</b> (найдено: {len(items)})', parse_mode='HTML')

    for hw_row in items:
        hw_id, hw_date, hw_subj, hw_text, hw_photo, hw_doc_id, hw_doc_name, hw_queue_id = hw_row[:8]
        photo_ids = parse_photo_ids(hw_photo)
        item_text, markup = render_hw_card_payload(hw_row, user_obj, target_group_name=target_group_name)

        try:
            send_hw_media(
                chat_id,
                photo_ids,
                item_text,
                reply_markup=markup,
                doc_file_id=hw_doc_id,
                doc_file_name=hw_doc_name,
            )
        except Exception:
            bot.send_message(
                chat_id,
                item_text + '\n<i>(не удалось загрузить вложение)</i>',
                reply_markup=markup,
                parse_mode='HTML',
            )


def send_homeworks_for_date_view(chat_id, user_obj, group_name, date_str):
    with db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            'SELECT id, target_date, subject_name, hw_text, photo_file_id, doc_file_id, doc_file_name, queue_id, added_by '
            'FROM homeworks '
            'WHERE group_name = ? AND target_date = ? ORDER BY id ASC',
            (group_name, date_str),
        )
        items = cursor.fetchall()
    send_hw_items(chat_id, user_obj, items, f'ДЗ на {date_str}', target_group_name=group_name)


def send_homeworks_for_period(chat_id, user_obj, group_name, days_limit=None):
    today = now_local().date()
    with db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            'SELECT id, target_date, subject_name, hw_text, photo_file_id, doc_file_id, doc_file_name, queue_id, added_by '
            'FROM homeworks '
            'WHERE group_name = ?',
            (group_name,),
        )
        all_items = cursor.fetchall()

    filtered = []
    for item in all_items:
        try:
            d = datetime.strptime(item[1], '%d.%m.%Y').date()
        except ValueError:
            continue
        if d < today:
            continue
        if days_limit is not None and d > today + timedelta(days=days_limit):
            continue
        filtered.append((d, item))

    filtered.sort(key=lambda x: (x[0], x[1][0]))
    items = [x[1] for x in filtered]
    title = 'ДЗ на ближайшую неделю' if days_limit == 7 else 'Все актуальные ДЗ'
    send_hw_items(chat_id, user_obj, items, title, target_group_name=group_name)


# --- ИНЛАЙН-РЕЖИМ (@бот в любом чате) ---
@bot.inline_handler(func=lambda query: True)
def inline_schedule_handler(inline_query):
    user_id = inline_query.from_user.id
    user = get_user_data(user_id)
    results = []

    if not user or not is_group_valid(user[7], user[0], user[1]):
        results.append(
            InlineQueryResultArticle(
                id='need_reg',
                title='⚠️ Сначала выбери факультет и группу в боте',
                description='Открой бота и нажми /start, чтобы привязать свою группу',
                input_message_content=InputTextMessageContent(
                    '⚠️ Чтобы делиться расписанием и ДЗ, сначала нужно выбрать свой факультет и группу в боте (/start).'
                ),
            )
        )
        bot.answer_inline_query(inline_query.id, results, cache_time=5, is_personal=True)
        return

    course, group_name, fac_code = user[0], user[1], user[7]
    q = (inline_query.query or '').strip().lower()
    today = now_local()

    if not q or any(w in q for w in ['сегодня', 'today', 'пары', 'расп']):
        sched_today = get_schedule_text(today, course, group_name, faculty_code=fac_code)
        results.append(
            InlineQueryResultArticle(
                id=f'today_{today.strftime("%Y%m%d")}',
                title=f'📅 Расписание на сегодня ({today.strftime("%d.%m")})',
                description=f'{get_faculty_short(fac_code)} • {group_name}',
                input_message_content=InputTextMessageContent(sched_today, parse_mode='HTML'),
            )
        )

    if not q or any(w in q for w in ['завтра', 'tomorrow', 'пары', 'расп']):
        tomorrow = today + timedelta(days=1)
        sched_tomorrow = get_schedule_text(tomorrow, course, group_name, faculty_code=fac_code)
        results.append(
            InlineQueryResultArticle(
                id=f'tomorrow_{tomorrow.strftime("%Y%m%d")}',
                title=f'⏭ Расписание на завтра ({tomorrow.strftime("%d.%m")})',
                description=f'{get_faculty_short(fac_code)} • {group_name}',
                input_message_content=InputTextMessageContent(sched_tomorrow, parse_mode='HTML'),
            )
        )

    if not q or any(w in q for w in ['дз', 'домаш', 'hw', 'задан']):
        with db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                'SELECT id, target_date, subject_name, hw_text, photo_file_id, doc_file_id, doc_file_name '
                'FROM homeworks WHERE group_name = ?',
                (group_name,),
            )
            all_hws = cursor.fetchall()

        today_d = today.date()
        upcoming = []
        for hw_id, h_date, h_subj, h_text, h_photo, h_doc_id, h_doc_name in all_hws:
            try:
                d_obj = datetime.strptime(h_date, '%d.%m.%Y').date()
                if d_obj >= today_d:
                    upcoming.append((d_obj, hw_id, h_date, h_subj, h_text, h_photo, h_doc_id, h_doc_name))
            except ValueError:
                pass
        upcoming.sort(key=lambda x: (x[0], x[1]))

        bot_uname = get_bot_username()

        if upcoming:
            hw_msg = f'📚 <b>АКТУАЛЬНЫЕ ДЗ — {esc(group_name)}</b>\n\n'
            has_any_media = False
            for _, hw_id, h_date, h_subj, h_text, h_photo, h_doc_id, h_doc_name in upcoming[:10]:
                p_list = parse_photo_ids(h_photo)
                p_cnt = len(p_list)
                badges = []
                if p_cnt:
                    badges.append(f'📷 {p_cnt} фото')
                    has_any_media = True
                if h_doc_id:
                    badges.append(f'📎 {esc(h_doc_name or "файл")}')
                    has_any_media = True
                p_badge = f' [{", ".join(badges)}]' if badges else ''
                desc = esc(h_text) if h_text else '<i>(см. вложение ниже или в боте)</i>'
                hw_msg += f'📅 <b>{h_date}</b> — <b>{esc(h_subj)}</b>{p_badge}\n📝 {desc}\n\n'
            hw_preview = f'Заданий: {len(upcoming)} (ближайшее: {upcoming[0][3]})'
            summary_kb = None
            if has_any_media and bot_uname:
                summary_kb = InlineKeyboardMarkup()
                summary_kb.add(
                    InlineKeyboardButton(
                        '📎 Открыть все вложения в боте',
                        url=f'https://t.me/{bot_uname}?start=hw_all',
                    )
                )
        else:
            hw_msg = f'📚 <b>ДОМАШНИЕ ЗАДАНИЯ — {esc(group_name)}</b>\n\n✨ Актуальных заданий пока нет!'
            hw_preview = 'Актуальных заданий пока нет'
            summary_kb = None

        results.append(
            InlineQueryResultArticle(
                id=f'hw_summary_{today.strftime("%Y%m%d")}',
                title='📚 Все актуальные ДЗ (списком)',
                description=hw_preview,
                input_message_content=InputTextMessageContent(hw_msg.strip(), parse_mode='HTML'),
                reply_markup=summary_kb,
            )
        )

        # Отдельные карточки по каждому ДЗ (с реальным фото / документом, если есть вложение!)
        for _, hw_id, h_date, h_subj, h_text, h_photo, h_doc_id, h_doc_name in upcoming[:15]:
            p_list = parse_photo_ids(h_photo)
            extra_notes = []
            if len(p_list) > 1:
                extra_notes.append(f'📷 Прикреплено фото: {len(p_list)} шт. (в чате отображается 1-е)')
            if p_list and h_doc_id:
                extra_notes.append(f'📎 Также прикреплён файл: {esc(h_doc_name or "документ")}')
            extra_footer = ('\n\n<i>' + '\n'.join(extra_notes) + '</i>') if extra_notes else ''

            card_caption = (
                f'📚 <b>ДОМАШНЕЕ ЗАДАНИЕ ({esc(group_name)})</b>\n'
                f'📅 <b>Дедлайн:</b> {h_date}\n'
                f'📘 <b>{esc(h_subj)}</b>\n\n'
                f'📝 {esc(h_text) if h_text else "<i>(см. прикреплённое вложение)</i>"}'
                f'{extra_footer}'
            )

            item_kb = None
            if (len(p_list) > 1 or (p_list and h_doc_id)) and bot_uname:
                item_kb = InlineKeyboardMarkup()
                item_kb.add(
                    InlineKeyboardButton(
                        '📎 Все фото и файлы в боте',
                        url=f'https://t.me/{bot_uname}?start=hw_item_{hw_id}',
                    )
                )

            short_desc = (h_text[:70] + '…' if len(h_text) > 70 else h_text) if h_text else 'Вложение к заданию'

            if p_list:
                results.append(
                    InlineQueryResultCachedPhoto(
                        id=f'hw_photo_{hw_id}',
                        photo_file_id=p_list[0],
                        title=f'📷 ДЗ ({h_date}): {h_subj}',
                        description=short_desc,
                        caption=card_caption[:1024],
                        parse_mode='HTML',
                        reply_markup=item_kb,
                    )
                )
            elif h_doc_id:
                results.append(
                    InlineQueryResultCachedDocument(
                        id=f'hw_doc_{hw_id}',
                        title=f'📎 ДЗ ({h_date}): {h_subj}',
                        document_file_id=h_doc_id,
                        description=f'{h_doc_name or "Файл"} • {short_desc}',
                        caption=card_caption[:1024],
                        parse_mode='HTML',
                        reply_markup=item_kb,
                    )
                )
            else:
                results.append(
                    InlineQueryResultArticle(
                        id=f'hw_text_{hw_id}',
                        title=f'📝 ДЗ ({h_date}): {h_subj}',
                        description=short_desc,
                        input_message_content=InputTextMessageContent(card_caption, parse_mode='HTML'),
                    )
                )

    bot.answer_inline_query(inline_query.id, results, cache_time=10, is_personal=True)


# --- ОБРАБОТЧИКИ МЕНЮ И КНОПОК ---
_BOT_USERNAME_CACHE = None


def get_bot_username():
    global _BOT_USERNAME_CACHE
    if _BOT_USERNAME_CACHE:
        return _BOT_USERNAME_CACHE
    try:
        me = bot.get_me()
        if me and me.username:
            _BOT_USERNAME_CACHE = me.username
            return _BOT_USERNAME_CACHE
    except Exception:
        pass
    return ''


@bot.message_handler(commands=['start'], chat_types=['private'])
def cmd_start(message):
    sync_username(message.from_user)
    clear_user_state(message.chat.id)
    start_parts = (message.text or '').split(maxsplit=1)
    deep_arg = start_parts[1].strip() if len(start_parts) > 1 else ''
    admin_flag = is_admin_user(message.from_user)
    user = get_user_data(message.from_user.id)
    if user is None:
        if admin_flag:
            uname = (message.from_user.username or '').lower()
            with db_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    'INSERT OR IGNORE INTO users (user_id, faculty, course, group_name, subscribed, username, is_starosta, is_hw_manager, notify_updates, managed_group) '
                    'VALUES (?, "ЕМФ", 1, "", 0, ?, 1, 1, 0, "")',
                    (message.from_user.id, uname),
                )
            invalidate_user_cache(message.from_user.id)
            user = get_user_data(message.from_user.id)
        else:
            start_onboarding(message)
            return

    course, group_name, subscribed, _, _, notify_updates, managed_group, fac_code = user
    has_valid_group = is_group_valid(fac_code, course, group_name)
    if not has_valid_group and not admin_flag:
        start_onboarding(message)
        return
    starosta_flag = is_user_starosta(message.from_user)
    hw_flag = can_manage_hw(message.from_user)
    sub_status = 'включена (07:00 и 18:00)' if subscribed else 'выключена'
    notify_status = 'включены' if notify_updates else 'выключены'
    if admin_flag:
        role_text = '👑 Администратор'
    elif starosta_flag:
        role_text = '🌟 Староста'
    else:
        role_text = 'Студент'

    group_info_block = (
        f'🏛 <b>Факультет:</b> {esc(get_faculty_short(fac_code))} ({esc(get_faculty_full(fac_code))})\n'
        f'🎓 <b>Курс:</b> {course}\n'
        f'📚 <b>Направление:</b> {esc(group_name)}\n'
        if has_valid_group
        else '📚 <b>Направление:</b> <i>не выбрано (загрузи расписание в Спецвозможностях)</i>\n'
    )

    if deep_arg == 'hw_all' and has_valid_group:
        send_homeworks_for_period(message.chat.id, message.from_user, group_name, days_limit=None)
        return

    if deep_arg.startswith('hw_item_') and has_valid_group:
        try:
            target_hw_id = int(deep_arg.replace('hw_item_', ''))
            with db_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    'SELECT id, target_date, subject_name, hw_text, photo_file_id, doc_file_id, doc_file_name, queue_id, added_by '
                    'FROM homeworks WHERE id = ? AND group_name = ?',
                    (target_hw_id, group_name),
                )
                hw_single = cursor.fetchone()
            if hw_single:
                send_hw_items(
                    message.chat.id,
                    message.from_user,
                    [hw_single],
                    f'Вложения к ДЗ ({hw_single[1]})',
                    target_group_name=group_name,
                )
                return
        except Exception:
            pass

    bot.send_message(
        message.chat.id,
        f'👋 <b>С возвращением!</b> Бот готов к работе.\n\n'
        f'<blockquote>{group_info_block}'
        f'👤 <b>Статус:</b> {role_text}\n'
        f'🗓 <b>Авторассылка:</b> {sub_status}\n'
        f'🔔 <b>Уведомления (пары и ДЗ):</b> {notify_status}</blockquote>\n\n'
        f'👇 <i>Выбери нужное действие в меню ниже:</i>',
        reply_markup=get_main_keyboard(bool(subscribed), starosta_flag, hw_flag),
        parse_mode='HTML',
    )


@bot.message_handler(
    func=lambda msg: msg.text in [
        '📅 Сегодня',
        '🗓 Неделя',
        '📚 Домашка',
        '🔍 Поиск',
        '⚙️ Настройки',
        '🛠 Спецвозможности',
        'Расписание на сегодня',
        '📅 Расписание на сегодня',
        'Расписание на неделю',
        '🗓 Расписание на неделю',
        '📚 Домашние задания',
        'Домашние задания',
        'Сменить группу',
        '⚙️ Сменить группу',
        '⚙️ Изменить группу',
        'Включить рассылку (07:00)',
        '🔔 Включить рассылку (07:00 и 18:00)',
        '🔔 Включить рассылку в 7:00',
        'Отключить утреннюю рассылку',
        '🔕 Отключить рассылку',
        'Перезапустить бота (/start)',
        '🔄 Перезапустить бота (/start)',
        '➕ Добавить замену',
        'Добавить замену',
        '🗑 Удалить пару',
        'Удалить пару',
        '📋 Активные замены/отмены',
        '📚 Добавить ДЗ',
    ]
)
def handle_menu(message):
    sync_username(message.from_user)
    clear_user_state(message.chat.id, keep_payload=True)

    user_id = message.from_user.id
    text_lower = message.text.lower()

    if message.text == '🔍 Поиск':
        open_search_menu(message.chat.id)
        return

    if message.text == '⚙️ Настройки':
        open_settings_menu(message.chat.id, message.from_user)
        return

    if message.text == '🛠 Спецвозможности':
        open_special_menu(message.chat.id, message.from_user)
        return

    if 'Перезапустить бота' in message.text:
        cmd_start(message)
        return

    if 'Сменить группу' in message.text or 'Изменить группу' in message.text:
        start_onboarding(message)
        return

    if 'Добавить замену' in message.text:
        start_add_replacement(message)
        return

    if 'Удалить пару' in message.text:
        start_delete_pair(message)
        return

    if 'Активные замены/отмены' in message.text:
        show_active_changes(message)
        return

    if 'Добавить ДЗ' in message.text:
        start_add_homework(message)
        return

    if 'Домашние задания' in message.text or 'Домашка' in message.text:
        open_homework_menu(message)
        return

    user = get_user_data(user_id)
    if not user or not is_group_valid(user[7], user[0], user[1]):
        start_onboarding(message)
        return

    course, group_name, fac_code = user[0], user[1], user[7]

    if 'рассылку' in message.text:
        open_settings_menu(message.chat.id, message.from_user)
        return

    if 'сегодня' in text_lower:
        text = get_schedule_text(now_local(), course, group_name, faculty_code=fac_code)
        bot.send_message(
            user_id,
            text,
            reply_markup=get_pagination_keyboard(0, group_name),
            parse_mode='HTML',
        )
        return

    if 'недел' in text_lower:
        send_week_schedule_message(
            user_id, now_local(), course, group_name, week_offset=0, faculty_code=fac_code
        )
        return


@bot.callback_query_handler(func=lambda call: True)
def callback_router(call):
    sync_username(call.from_user)
    user_id = call.from_user.id
    chat_id = call.message.chat.id
    msg_id = call.message.message_id
    data = call.data

    if not (
        data in ('settings_toggle_notify', 'settings_toggle_sub')
        or data.startswith('hwq_')
    ):
        try:
            bot.answer_callback_query(call.id)
        except Exception:
            pass

    if data == 'cancel_action':
        clear_user_state(chat_id)
        try:
            bot.edit_message_text('❌ Действие отменено.', chat_id, msg_id)
        except Exception:
            pass
        return

    # --- ВЫБОР ФАКУЛЬТЕТА -> КУРСА -> НАПРАВЛЕНИЯ ---
    if data == 'back_to_faculties':
        clear_user_state(user_id)
        try:
            bot.edit_message_text(
                '🏛 <b>Выбор факультета</b>\n\n👇 Выбери свой <b>факультет</b>:',
                chat_id,
                msg_id,
                reply_markup=build_faculty_keyboard(),
                parse_mode='HTML',
            )
        except Exception:
            pass
        return

    if data.startswith('sel_fac_'):
        fac_code = data.replace('sel_fac_', '')
        meta = FACULTIES_META.get(fac_code)
        if not meta:
            return
        courses = FACULTY_STRUCTURE.get(fac_code, {})
        has_any_groups = any(bool(grps) for grps in courses.values())

        if not has_any_groups:
            markup = InlineKeyboardMarkup(row_width=1)
            markup.add(InlineKeyboardButton('🔙 Выбрать другой факультет', callback_data='back_to_faculties'))
            try:
                bot.edit_message_text(
                    f'{meta["emoji"]} <b>{esc(meta["full"])} ({esc(meta["short"])})</b>\n\n'
                    f'⏳ <i>Расписание для этого факультета ещё не загружено и появится в ближайшее время!</i>\n'
                    f'Пока ты можешь выбрать другой факультет.',
                    chat_id,
                    msg_id,
                    reply_markup=markup,
                    parse_mode='HTML',
                )
            except Exception:
                pass
            return

        set_user_state(user_id, '', {'faculty': fac_code})
        try:
            bot.edit_message_text(
                f'{meta["emoji"]} Выбран факультет: <b>{esc(meta["full"])} ({esc(meta["short"])})</b>.\n\n'
                f'👇 Теперь выбери свой <b>курс</b>:',
                chat_id,
                msg_id,
                reply_markup=build_course_keyboard(fac_code),
                parse_mode='HTML',
            )
        except Exception:
            pass
        return

    if data.startswith('sel_course_'):
        parts = data.replace('sel_course_', '').split('_')
        fac_code = parts[0]
        course = int(parts[1])
        groups = FACULTY_STRUCTURE.get(fac_code, {}).get(course, [])
        if not groups:
            return

        markup = InlineKeyboardMarkup(row_width=1)
        for idx, g_name in enumerate(groups):
            markup.add(
                InlineKeyboardButton(f'📘 {g_name}', callback_data=f'sel_group_{fac_code}_{course}_{idx}')
            )
        markup.add(InlineKeyboardButton('🔙 Сменить курс', callback_data=f'sel_fac_{fac_code}'))

        set_user_state(user_id, '', {
            'faculty': fac_code,
            'course': course,
            'current_groups': groups,
        })
        try:
            bot.edit_message_text(
                f'🏛 <b>{esc(get_faculty_short(fac_code))}</b>  •  ✅ Выбран <b>{course} курс</b>.\n\n'
                f'👇 Теперь укажи своё <b>направление подготовки</b>:',
                chat_id,
                msg_id,
                reply_markup=markup,
                parse_mode='HTML',
            )
        except Exception:
            pass
        return

    if data.startswith('sel_group_'):
        parts = data.replace('sel_group_', '').split('_')
        fac_code = parts[0]
        course = int(parts[1])
        g_idx = int(parts[2])

        _, state = get_user_state(user_id)
        groups_list = state.get('current_groups') or FACULTY_STRUCTURE.get(fac_code, {}).get(course, [])
        if not groups_list:
            return
        group_name = groups_list[g_idx] if g_idx < len(groups_list) else groups_list[0]
        clear_user_state(user_id)
        uname = (call.from_user.username or '').lower()
        is_admin = 1 if uname == ADMIN_USERNAME.lower() else 0

        with db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                'SELECT subscribed, is_starosta, is_hw_manager, notify_updates, managed_group FROM users WHERE user_id = ?',
                (user_id,),
            )
            existing = cursor.fetchone()
            sub_val = existing[0] if existing else 0
            starosta_val = 1 if (is_admin or (existing and existing[1] == 1)) else 0
            hw_val = 1 if (is_admin or (existing and existing[2] == 1)) else 0
            notify_val = existing[3] if existing else 0
            # Восстанавливаем managed_group старосте после ежегодного сброса курсов, но не даём перезаписать уже закреплённую группу
            if existing and existing[4]:
                managed_val = existing[4]
            elif starosta_val == 1 and not is_admin:
                managed_val = group_name
            else:
                managed_val = ''

            cursor.execute(
                'INSERT OR REPLACE INTO users (user_id, faculty, course, group_name, subscribed, username, is_starosta, is_hw_manager, notify_updates, managed_group) '
                'VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                (user_id, fac_code, course, group_name, sub_val, uname, starosta_val, hw_val, notify_val, managed_val),
            )
        invalidate_user_cache(user_id)

        try:
            bot.delete_message(chat_id, msg_id)
        except Exception:
            pass
        bot.send_message(
            user_id,
            f'🎉 <b>Профиль успешно сохранён!</b>\n\n'
            f'<blockquote>🏛 <b>Факультет:</b> {esc(get_faculty_short(fac_code))} ({esc(get_faculty_full(fac_code))})\n'
            f'🎓 <b>Курс:</b> {course}\n'
            f'📚 <b>Направление:</b> {esc(group_name)}</blockquote>\n\n'
            'Расписание автоматически подстраивается под текущую учебную неделю.\n'
            '🔕 <i>Уведомления выключены по умолчанию. Включить оповещения об изменении пар и ДЗ можно в разделе <b>⚙️ Настройки</b>.</i>',
            reply_markup=get_main_keyboard(
                subscribed=bool(sub_val),
                is_starosta=bool(starosta_val),
                is_hw_manager=bool(hw_val),
            ),
            parse_mode='HTML',
        )
        return

    # --- ИМПОРТ РАСПИСАНИЯ ИЗ .DOCX / .DOC ---
    if data.startswith('docx_setfac_'):
        if not is_admin_user(call.from_user):
            return
        rest = data.replace('docx_setfac_', '')
        fac_code, token = rest.split('_', 1)
        payload = get_pending_docx_import(token)
        if not payload:
            bot.send_message(chat_id, '⚠️ Сессия импорта устарела. Отправь файл заново.')
            return
        payload['faculty_code'] = fac_code
        save_pending_docx_import(token, payload)
        year_start_str = payload['year_start_dt'].strftime('%d.%m.%Y')
        delta_prev = (now_local().date() - payload['year_start_dt'].date()).days
        cur_week = 1 if delta_prev < 0 else (delta_prev // 7) + 1
        by_group = {}
        for r in payload['records']:
            by_group[(r['course'], r['group'])] = by_group.get((r['course'], r['group']), 0) + 1
        preview_lines = [
            f'• {esc(grp)}: <b>{cnt}</b> записей'
            for (crs, grp), cnt in sorted(by_group.items())
        ]
        text = (
            f'✅ <b>Файл расписания готов к загрузке!</b>\n\n'
            f'🏛 <b>Выбран факультет:</b> {esc(get_faculty_short(fac_code))} ({esc(get_faculty_full(fac_code))})\n'
            f'📊 <b>Всего строк:</b> {len(payload["records"])}\n\n'
            f'🎓 <b>Распознанные группы:</b>\n' + '\n'.join(preview_lines) +
            f'\n\n📆 <b>Счётчик учебных недель:</b> без сброса от <code>{year_start_str}</code> (сейчас <b>{cur_week}-я</b> неделя)'
        )
        markup = InlineKeyboardMarkup(row_width=2)
        fac_btns = []
        for f_code in FACULTY_ORDER:
            mark = '✅ ' if f_code == fac_code else ''
            fac_btns.append(
                InlineKeyboardButton(
                    f'{mark}{get_faculty_short(f_code)}',
                    callback_data=f'docx_setfac_{f_code}_{token}',
                )
            )
        markup.add(*fac_btns)
        markup.add(
            InlineKeyboardButton(
                f'🚀 Применить ({get_faculty_short(fac_code)}, старт {year_start_str})',
                callback_data=f'docx_apply_auto_{token}',
            )
        )
        markup.add(
            InlineKeyboardButton(
                '📅 Изменить дату 1-й недели учебного года',
                callback_data=f'docx_apply_manual_{token}',
            )
        )
        markup.add(InlineKeyboardButton('❌ Отмена', callback_data=f'docx_cancel_{token}'))
        try:
            bot.edit_message_text(text, chat_id, msg_id, reply_markup=markup, parse_mode='HTML')
        except Exception:
            pass
        return

    if data.startswith('docx_cancel_'):
        if not is_admin_user(call.from_user):
            return
        clear_user_state(chat_id)
        token = data.replace('docx_cancel_', '')
        get_pending_docx_import(token, pop=True)
        try:
            bot.edit_message_text('❌ Импорт нового расписания отменён.', chat_id, msg_id)
        except Exception:
            pass
        return

    if data.startswith('docx_apply_auto_'):
        if not is_admin_user(call.from_user):
            return
        token = data.replace('docx_apply_auto_', '')
        payload = get_pending_docx_import(token)
        if not payload:
            bot.send_message(chat_id, '⚠️ Сессия импорта устарела. Отправь файл заново.')
            return
        try:
            bot.delete_message(chat_id, msg_id)
        except Exception:
            pass
        apply_parsed_docx_schedule(chat_id, user_id, token, payload['year_start_dt'])
        return

    if data.startswith('docx_apply_manual_'):
        if not is_admin_user(call.from_user):
            return
        token = data.replace('docx_apply_manual_', '')
        if not get_pending_docx_import(token):
            bot.send_message(chat_id, '⚠️ Сессия импорта устарела. Отправь файл заново.')
            return
        clear_user_state(chat_id)
        set_user_state(chat_id, 'awaiting_docx_date', {'import_token': token})
        default_start_hint = get_default_academic_year_start().strftime('%d.%m.%Y')
        bot.edit_message_text(
            f'📅 Введи дату <b>понедельника 1-й учебной недели учебного года</b> в формате <code>ДД.ММ.ГГГГ</code>\n'
            f'<i>(по умолчанию: {default_start_hint})</i>\n'
            f'<i>Для отмены напиши «Отмена».</i>',
            chat_id,
            msg_id,
            parse_mode='HTML',
        )
        return

    # --- ПОИСК ПО ПРЕПОДАВАТЕЛЮ И СВОБОДНЫМ АУДИТОРИЯМ ---
    if data == 'search_back':
        open_search_menu(chat_id, msg_id)
        return

    if data == 'search_teacher':
        clear_user_state(chat_id, keep_payload=True)
        show_teacher_period_choice(chat_id, msg_id)
        return

    if data.startswith('tsearch_period_'):
        clear_user_state(chat_id, keep_payload=True)
        period = data.replace('tsearch_period_', '')
        period_labels = {
            'today': 'на сегодня',
            'tomorrow': 'на завтра',
            'week': 'на всю текущую неделю',
        }
        teachers_sample = get_all_unique_teachers()[:5]
        hint_line = (
            f'<i>Например: {", ".join(esc(t) for t in teachers_sample)}</i>\n'
            if teachers_sample
            else ''
        )
        set_user_state(chat_id, 'awaiting_teacher_query', {'teacher_period': period})
        bot.edit_message_text(
            f'👨‍🏫 <b>Поиск преподавателя ({period_labels.get(period, "")})</b>\n\n'
            f'<b>Шаг 2 из 2.</b> Введи фамилию или инициалы (например, <code>Ситнова И.А.</code>):\n'
            f'{hint_line}'
            f'<i>Для отмены напиши «Отмена».</i>',
            chat_id,
            msg_id,
            parse_mode='HTML',
        )
        return

    if data == 'tsearch_pick_all':
        _, state = get_user_state(chat_id)
        if not state or 'teacher_raw_query' not in state:
            try:
                bot.answer_callback_query(call.id, '⚠️ Список устарел, повтори поиск')
            except Exception:
                pass
            return
        try:
            bot.delete_message(chat_id, msg_id)
        except Exception:
            pass
        execute_teacher_search_for_period(
            chat_id, state['teacher_raw_query'], state.get('teacher_period', 'today')
        )
        return

    if data.startswith('tsearch_pick_'):
        _, state = get_user_state(chat_id)
        if not state or 'teacher_suggestions' not in state:
            try:
                bot.answer_callback_query(call.id, '⚠️ Список устарел, повтори поиск')
            except Exception:
                pass
            return
        idx = int(data.replace('tsearch_pick_', ''))
        suggestions = state['teacher_suggestions']
        if idx >= len(suggestions):
            return
        try:
            bot.delete_message(chat_id, msg_id)
        except Exception:
            pass
        execute_teacher_search_for_period(
            chat_id, suggestions[idx], state.get('teacher_period', 'today')
        )
        return

    if data.startswith('tsearch_re_'):
        period = data.replace('tsearch_re_', '')
        _, state = get_user_state(chat_id)
        if not state or 'last_teacher_query' not in state:
            show_teacher_period_choice(chat_id, msg_id)
            return
        execute_teacher_search_for_period(chat_id, state['last_teacher_query'], period)
        return

    if data == 'search_rooms':
        clear_user_state(chat_id, keep_payload=True)
        show_free_rooms_day_choice(chat_id, msg_id)
        return

    if data.startswith('freerooms_day_'):
        day_offset = int(data.replace('freerooms_day_', ''))
        show_free_rooms_pair_choice(chat_id, msg_id, day_offset=day_offset)
        return

    if data.startswith('freerooms_pair_'):
        parts = data.replace('freerooms_pair_', '').split('_')
        day_offset = int(parts[0])
        pair_num = int(parts[1])
        show_free_rooms_result(chat_id, msg_id, pair_num, day_offset=day_offset)
        return

    # --- НАСТРОЙКИ И ПРОФИЛЬ ---
    if data == 'settings_toggle_sub':
        user = get_user_data(user_id)
        if not user:
            try:
                bot.answer_callback_query(call.id)
            except Exception:
                pass
            return
        new_sub = 0 if user[2] == 1 else 1
        with db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                'UPDATE users SET subscribed = ? WHERE user_id = ?',
                (new_sub, user_id),
            )
        invalidate_user_cache(user_id)
        try:
            alert = (
                '🔔 Авторассылка включена (07:00 и 18:00 с ДЗ)'
                if new_sub
                else '🔕 Авторассылка отключена'
            )
            bot.answer_callback_query(call.id, alert)
        except Exception:
            pass
        open_settings_menu(chat_id, call.from_user, msg_id)
        return

    if data == 'settings_toggle_notify':
        user = get_user_data(user_id)
        if not user:
            try:
                bot.answer_callback_query(call.id)
            except Exception:
                pass
            return
        new_val = 0 if user[5] == 1 else 1
        with db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                'UPDATE users SET notify_updates = ? WHERE user_id = ?',
                (new_val, user_id),
            )
        invalidate_user_cache(user_id)
        try:
            alert = (
                '🔔 Уведомления о парах и ДЗ включены'
                if new_val
                else '🔕 Уведомления о парах и ДЗ выключены'
            )
            bot.answer_callback_query(call.id, alert)
        except Exception:
            pass
        open_settings_menu(chat_id, call.from_user, msg_id)
        return

    if data == 'settings_change_group':
        clear_user_state(chat_id)
        try:
            bot.edit_message_text(
                '🏛 <b>Смена факультета и группы</b>\n\n👇 Выбери свой <b>факультет</b>:',
                chat_id,
                msg_id,
                reply_markup=build_faculty_keyboard(),
                parse_mode='HTML',
            )
        except Exception:
            pass
        return

    if data == 'settings_restart':
        clear_user_state(chat_id)
        try:
            bot.delete_message(chat_id, msg_id)
        except Exception:
            pass
        user = get_user_data(user_id)
        if not user or not is_group_valid(user[7], user[0], user[1]):
            bot.send_message(
                chat_id,
                '👇 Выбери свой <b>факультет</b>:',
                reply_markup=build_faculty_keyboard(),
                parse_mode='HTML',
            )
            return
        starosta_flag = is_user_starosta(call.from_user)
        hw_flag = can_manage_hw(call.from_user)
        bot.send_message(
            chat_id,
            '🔄 <b>Главное меню обновлено!</b>',
            reply_markup=get_main_keyboard(bool(user[2]), starosta_flag, hw_flag),
            parse_mode='HTML',
        )
        return

    # --- СПЕЦВОЗМОЖНОСТИ ---
    if data == 'spec_back':
        clear_user_state(chat_id)
        open_special_menu(chat_id, call.from_user, msg_id)
        return

    if data == 'spec_add_repl':
        open_add_replacement(chat_id, call.from_user, msg_id)
        return

    if data == 'spec_del_pair':
        open_delete_pair(chat_id, call.from_user, msg_id)
        return

    if data == 'spec_active_changes':
        if not is_user_starosta(call.from_user):
            return
        _, group_name = get_managed_context(call.from_user)
        if not group_name:
            return
        text, markup = build_active_changes_view(group_name)
        markup.add(InlineKeyboardButton('🔙 Назад в спецвозможности', callback_data='spec_back'))
        try:
            bot.edit_message_text(text, chat_id, msg_id, reply_markup=markup, parse_mode='HTML')
        except Exception:
            bot.send_message(chat_id, text, reply_markup=markup, parse_mode='HTML')
        return

    if data == 'spec_add_hw':
        open_add_homework(chat_id, call.from_user, msg_id)
        return

    if data == 'hw_queue_no':
        if not can_manage_hw(call.from_user):
            return
        _, state = get_user_state(chat_id)
        draft = (state or {}).get('pending_hw_draft')
        if not draft:
            bot.send_message(chat_id, '⚠️ Сессия создания ДЗ устарела. Начни заново.')
            return
        try:
            bot.delete_message(chat_id, msg_id)
        except Exception:
            pass
        save_and_broadcast_homework(
            chat_id,
            draft['author_id'],
            draft['group_name'],
            draft['date_str'],
            draft['subject_name'],
            draft['hw_text'],
            draft['photo_ids'],
            doc_file_id=draft['doc_file_id'],
            doc_file_name=draft['doc_file_name'],
            queue_slots_count=None,
            author_label=draft.get('author_label') or format_hw_author_label(call.from_user),
        )
        return

    if data == 'hw_queue_yes':
        if not can_manage_hw(call.from_user):
            return
        _, state = get_user_state(chat_id)
        draft = (state or {}).get('pending_hw_draft')
        if not draft:
            bot.send_message(chat_id, '⚠️ Сессия создания ДЗ устарела. Начни заново.')
            return
        set_user_state(chat_id, 'awaiting_hw_slots_count', state)
        bot.edit_message_text(
            f'🔢 <b>Укажи количество вопросов / слотов для ДЗ по предмету «{esc(draft["subject_name"])}»</b>\n\n'
            f'Введи число от <b>1</b> до <b>60</b> (например, <code>12</code> или <code>20</code>):\n'
            f'<i>Для отмены напиши «Отмена».</i>',
            chat_id,
            msg_id,
            parse_mode='HTML',
        )
        return

    # --- РАЗБОР И ОСВОБОЖДЕНИЕ ВОПРОСОВ В ДОМАШНЕМ ЗАДАНИИ ---
    if data.startswith('hwq_pick_') or data.startswith('hwq_free_') or data.startswith('hwq_refresh_'):
        user_label = format_user_queue_label(call.from_user)
        cb_alert_text = ''
        cb_show_alert = False
        if data.startswith('hwq_pick_'):
            parts = data.replace('hwq_pick_', '').split('_')
            q_id = int(parts[0])
            chosen_slot = int(parts[1])
            u_data = get_user_data(user_id)
            user_group = u_data[1] if u_data else ''
            with db_connection() as conn:
                cursor = conn.cursor()
                cursor.execute('SELECT max_slots, is_open, group_name FROM queues WHERE id = ?', (q_id,))
                q_info = cursor.fetchone()
                if not q_info:
                    cb_alert_text = '⚠️ Очередь не найдена.'
                    cb_show_alert = True
                elif not is_admin_user(call.from_user) and user_group != q_info[2]:
                    cb_alert_text = '⛔ Занимать вопросы могут только студенты этой группы!'
                    cb_show_alert = True
                elif q_info[1] == 1 and 1 <= chosen_slot <= q_info[0]:
                    try:
                        cursor.execute(
                            'INSERT INTO queue_slots (queue_id, slot_num, user_id, user_label, joined_at) '
                            'VALUES (?, ?, ?, ?, ?)',
                            (q_id, chosen_slot, user_id, user_label, now_local().strftime('%d.%m.%Y %H:%M')),
                        )
                        cb_alert_text = f'✅ Вопрос №{chosen_slot} закреплён за тобой!'
                    except sqlite3.IntegrityError:
                        cb_alert_text = f'⚠️ Вопрос №{chosen_slot} только что занял кто-то другой!'
                        cb_show_alert = True

        elif data.startswith('hwq_free_'):
            parts = data.replace('hwq_free_', '').split('_')
            q_id = int(parts[0])
            target_slot = int(parts[1])
            with db_connection() as conn:
                cursor = conn.cursor()
                cursor.execute('SELECT group_name, created_by FROM queues WHERE id = ?', (q_id,))
                q_row = cursor.fetchone()
                q_grp = q_row[0] if q_row else ''
                q_creator = q_row[1] if q_row else None
                is_mod = can_edit_specific_hw(call.from_user, q_grp, q_creator)
                if is_mod:
                    cursor.execute(
                        'DELETE FROM queue_slots WHERE queue_id = ? AND slot_num = ?',
                        (q_id, target_slot),
                    )
                else:
                    cursor.execute(
                        'DELETE FROM queue_slots WHERE queue_id = ? AND slot_num = ? AND user_id = ?',
                        (q_id, target_slot, user_id),
                    )
                cb_alert_text = f'Слот №{target_slot} освобождён'
        else:
            q_id = int(data.replace('hwq_refresh_', ''))
            cb_alert_text = 'Список обновлён'

        try:
            bot.answer_callback_query(call.id, text=cb_alert_text or None, show_alert=cb_show_alert)
        except Exception:
            pass

        with db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                'SELECT id, target_date, subject_name, hw_text, photo_file_id, doc_file_id, doc_file_name, queue_id, added_by, group_name '
                'FROM homeworks WHERE queue_id = ?',
                (q_id,),
            )
            hw_full = cursor.fetchone()

        if hw_full:
            hw_row = hw_full[:9]
            grp_name = hw_full[9]
            new_text, new_markup = render_hw_card_payload(hw_row, call.from_user, target_group_name=grp_name)
            try:
                if call.message.content_type in ('photo', 'document'):
                    if len(new_text) <= 1024:
                        bot.edit_message_caption(
                            caption=new_text,
                            chat_id=chat_id,
                            message_id=msg_id,
                            reply_markup=new_markup,
                            parse_mode='HTML',
                        )
                    else:
                        bot.send_message(
                            chat_id,
                            new_text,
                            reply_markup=new_markup,
                            parse_mode='HTML',
                        )
                else:
                    bot.edit_message_text(
                        new_text,
                        chat_id=chat_id,
                        message_id=msg_id,
                        reply_markup=new_markup,
                        parse_mode='HTML',
                    )
            except ApiTelegramException as e_edit:
                if 'message is not modified' not in str(e_edit).lower():
                    try:
                        bot.send_message(
                            chat_id,
                            new_text,
                            reply_markup=new_markup,
                            parse_mode='HTML',
                        )
                    except Exception:
                        pass
            except Exception:
                pass
        return

    if data == 'spec_stats':
        send_stats_panel(chat_id, call.from_user)
        return

    if data == 'spec_broadcast':
        open_broadcast_menu(chat_id, call.from_user, msg_id)
        return

    if data == 'spec_upload_docx':
        if not is_admin_user(call.from_user):
            return
        markup = InlineKeyboardMarkup()
        markup.add(InlineKeyboardButton('🔙 Назад', callback_data='spec_back'))
        bot.edit_message_text(
            '📥 <b>ОБНОВЛЕНИЕ РАСПИСАНИЯ ИЗ WORD (.DOCX / .DOC)</b>\n\n'
            'Просто отправь мне в этот чат документ <code>.docx</code> (или <code>.doc</code>) с расписанием факультета (в табличной форме СИ УУНиТ).\n\n'
            '• Бот автоматически определит факультет (ЕМФ, Педфак, ФЭП, ТФ), курсы, направления, предметы, недели, преподавателей и аудитории.\n'
            '• Записи других факультетов в базе сохранятся нетронутыми.\n'
            '• Счётчик недель автоматически обновляется с началом нового учебного года (понедельник недели 1 сентября), а во 2-м семестре продолжает идти сквозным образом без сброса.\n'
            '• Перед перезаписью автоматически сделается резервная копия текущей БД.',
            chat_id,
            msg_id,
            reply_markup=markup,
            parse_mode='HTML',
        )
        return

    if data == 'spec_backup_now':
        if is_admin_user(call.from_user):
            send_db_backup(manual_chat_id=chat_id)
        return

    # --- НАВИГАЦИЯ ПО РАСПИСАНИЮ ---
    if data.startswith('date_'):
        offset = int(data.split('_')[1])
        user = get_user_data(user_id)
        if not user:
            return
        course, group_name, fac_code = user[0], user[1], user[7]
        text = get_schedule_text(
            now_local() + timedelta(days=offset), course, group_name, faculty_code=fac_code
        )
        try:
            bot.edit_message_text(
                text,
                chat_id,
                msg_id,
                reply_markup=get_pagination_keyboard(offset, group_name),
                parse_mode='HTML',
            )
        except Exception:
            pass
        return

    if data.startswith('week_'):
        week_offset = int(data.split('_')[1])
        user = get_user_data(user_id)
        if not user:
            return
        course, group_name, fac_code = user[0], user[1], user[7]
        target_date = now_local() + timedelta(days=week_offset * 7)
        send_week_schedule_message(
            chat_id, target_date, course, group_name, week_offset=week_offset, edit_msg_id=msg_id, faculty_code=fac_code
        )
        return

    # --- ЗАМЕНЫ И ОТМЕНЫ ЧЕРЕЗ INLINE-КНОПКИ ---
    if data.startswith('repl_date_') or data.startswith('del_date_'):
        if not is_user_starosta(call.from_user):
            return
        course, group_name = get_managed_context(call.from_user)
        if not group_name:
            return
        action = 'repl' if data.startswith('repl_') else 'del'
        date_str = data.split('_date_')[1]
        show_pair_selection_for_action(chat_id, msg_id, action, course, group_name, date_str)
        return

    if data in ('repl_back_date', 'del_back_date'):
        if not is_user_starosta(call.from_user):
            return
        _, group_name = get_managed_context(call.from_user)
        if not group_name:
            return
        action = 'repl' if data.startswith('repl_') else 'del'
        title = '📝 <b>Добавление замены</b>' if action == 'repl' else '🗑 <b>Отмена пары</b>'
        bot.edit_message_text(
            f'{title}\n<blockquote>🎓 {esc(group_name)}</blockquote>\n\nВыбери <b>дату</b>:',
            chat_id,
            msg_id,
            reply_markup=build_date_choice_markup(action),
            parse_mode='HTML',
        )
        return

    if data in ('repl_datemanual', 'del_datemanual'):
        if not is_user_starosta(call.from_user):
            return
        course, group_name = get_managed_context(call.from_user)
        if not group_name:
            return
        action = 'repl' if data.startswith('repl_') else 'del'
        clear_user_state(chat_id)
        set_user_state(chat_id, 'awaiting_manual_date', {
            'action': action,
            'course': course,
            'group_name': group_name,
        })
        try:
            bot.edit_message_text(
                f'⌨️ Введи дату в формате <code>ДД.ММ.ГГГГ</code> (например, <code>{now_local().strftime("%d.%m.%Y")}</code>)\n'
                f'<i>Для отмены напиши «Отмена».</i>',
                chat_id,
                msg_id,
                parse_mode='HTML',
            )
        except Exception:
            pass

        return

    if data.startswith('repl_pair_'):
        if not is_user_starosta(call.from_user):
            return
        _, group_name = get_managed_context(call.from_user)
        if not group_name:
            return
        parts = data.split('_')
        date_str = parts[2]
        pair_num = int(parts[3])
        time_slot = NUM_TO_TIME_SLOT.get(pair_num, str(pair_num))

        clear_user_state(chat_id)
        set_user_state(chat_id, 'awaiting_repl_subject', {
            'group_name': group_name,
            'date_str': date_str,
            'time_slot': time_slot,
        })
        bot.edit_message_text(
            f'✅ <b>Дата:</b> {date_str}\n'
            f'{format_pair_prefix(time_slot)}\n\n'
            f'<b>Шаг 3 из 3.</b> Введи <b>название предмета, преподавателя и аудиторию</b>\n'
            f'<i>(например: Экологический мониторинг (ПЗ) — ауд. 206, доц. Кужина Г.Ш.)</i>\n'
            f'<i>Для отмены напиши «Отмена».</i>',
            chat_id,
            msg_id,
            parse_mode='HTML',
        )
        return

    if data.startswith('del_pair_'):
        if not is_user_starosta(call.from_user):
            return
        course, group_name = get_managed_context(call.from_user)
        if not group_name:
            return
        parts = data.split('_')
        date_str = parts[2]
        pair_num = int(parts[3])
        time_slot = NUM_TO_TIME_SLOT.get(pair_num, str(pair_num))

        try:
            bot.delete_message(chat_id, msg_id)
        except Exception:
            pass
        execute_delete_pair(chat_id, user_id, course, group_name, date_str, time_slot)
        return

    # --- ОТКАТ АКТИВНЫХ ЗАМЕН И ОТМЕН ---
    if data.startswith('rev_repl_') or data.startswith('rev_canc_'):
        if not is_user_starosta(call.from_user):
            return
        _, group_name = get_managed_context(call.from_user)
        if not group_name:
            return
        rec_id = int(data.split('_')[2])
        notify_text = None

        with db_connection() as conn:
            cursor = conn.cursor()
            if data.startswith('rev_repl_'):
                cursor.execute(
                    'SELECT target_date, time_slot, subject_info FROM replacements '
                    'WHERE id = ? AND group_name = ?',
                    (rec_id, group_name),
                )
                row = cursor.fetchone()
                cursor.execute(
                    'DELETE FROM replacements WHERE id = ? AND group_name = ?',
                    (rec_id, group_name),
                )
                if row:
                    r_date, r_time, r_subj = row
                    notify_text = (
                        f'↩️ <b>ЗАМЕНА УДАЛЕНА ИЗ РАСПИСАНИЯ</b>\n'
                        f'<blockquote>🎓 <b>Группа:</b> {esc(group_name)}\n'
                        f'📅 <b>Дата:</b> {r_date}</blockquote>\n\n'
                        f'{format_pair_prefix(r_time)}\n'
                        f'<s>{esc(r_subj)}</s>'
                    )
            else:
                cursor.execute(
                    'SELECT target_date, time_slot FROM cancelled_lessons '
                    'WHERE id = ? AND group_name = ?',
                    (rec_id, group_name),
                )
                row = cursor.fetchone()
                cursor.execute(
                    'DELETE FROM cancelled_lessons WHERE id = ? AND group_name = ?',
                    (rec_id, group_name),
                )
                if row:
                    c_date, c_time = row
                    notify_text = (
                        f'↩️ <b>ОТМЕНА ПАРЫ СНЯТА (ПАРА ВОЗВРАЩЕНА)</b>\n'
                        f'<blockquote>🎓 <b>Группа:</b> {esc(group_name)}\n'
                        f'📅 <b>Дата:</b> {c_date}</blockquote>\n\n'
                        f'{format_pair_prefix(c_time)} — снова в расписании.'
                    )

            cursor.execute(
                'SELECT user_id FROM users WHERE group_name = ? AND notify_updates = 1',
                (group_name,),
            )
            group_users = cursor.fetchall()

        if notify_text:
            for (u_id,) in group_users:
                if u_id != user_id:
                    enqueue_message(u_id, notify_text, parse_mode='HTML')

        text, markup = build_active_changes_view(group_name)
        try:
            bot.edit_message_text(text, chat_id, msg_id, reply_markup=markup, parse_mode='HTML')
        except Exception:
            pass
        return

    # --- МАССОВАЯ РАССЫЛКА (/broadcast) ---
    if data.startswith('bc_target_'):
        if not is_admin_user(call.from_user):
            return
        clear_user_state(chat_id)
        target = data.replace('bc_target_', '')
        if target == 'all':
            target_label = 'всем пользователям'
        elif target.startswith('fac_'):
            target_label = f'факультету {get_faculty_short(target.replace("fac_", ""))}'
        else:
            target_label = f'{target} курсу'
        set_user_state(chat_id, 'awaiting_broadcast', {'target': target})
        bot.edit_message_text(
            f'📣 <b>Аудитория:</b> {target_label}\n\n'
            f'Теперь отправь текст объявления (можно прикрепить фото с подписью).\n'
            f'<i>Для отмены напиши «Отмена».</i>',
            chat_id,
            msg_id,
            parse_mode='HTML',
        )
        return

    # --- ДОМАШНИЕ ЗАДАНИЯ (ДОБАВЛЕНИЕ, ПРОСМОТР, УДАЛЕНИЕ) ---
    if data == 'hw_open_add':
        open_add_homework(chat_id, call.from_user, msg_id)
        return

    if data == 'hwadd_bydate_menu':
        if not can_manage_hw(call.from_user):
            return
        open_add_homework_by_date(chat_id, call.from_user, msg_id)
        return

    if data.startswith('hwadd_picksubj_'):
        if not can_manage_hw(call.from_user):
            return
        _, state = get_user_state(chat_id)
        entries = (state or {}).get('hw_subj_entries', [])
        idx = int(data.replace('hwadd_picksubj_', ''))
        if idx >= len(entries):
            return
        ent = entries[idx]
        group_name = state.get('hw_group', '')
        course = state.get('hw_course') or find_course_by_group(group_name)
        subj_name = ent['subject']
        up_dates = ent.get('upcoming_dates') or []
        try:
            bot.delete_message(chat_id, msg_id)
        except Exception:
            pass
        if up_dates:
            ask_hw_content(chat_id, group_name, up_dates[0], subj_name, upcoming_dates=up_dates)
        else:
            today = now_local()
            d1 = (today + timedelta(days=1)).strftime('%d.%m.%Y')
            d2 = (today + timedelta(days=2)).strftime('%d.%m.%Y')
            d7 = (today + timedelta(days=7)).strftime('%d.%m.%Y')
            set_user_state(chat_id, '', {
                'hw_group': group_name,
                'hw_course': course,
                'hw_pending_custom_subj': subj_name,
            })
            markup = InlineKeyboardMarkup(row_width=1)
            markup.add(
                InlineKeyboardButton(f'⏭ На завтра ({d1})', callback_data=f'hwadd_custdate_{d1}'),
                InlineKeyboardButton(f'⏩ На послезавтра ({d2})', callback_data=f'hwadd_custdate_{d2}'),
                InlineKeyboardButton(f'🗓 Через неделю ({d7})', callback_data=f'hwadd_custdate_{d7}'),
                InlineKeyboardButton('⌨️ Ввести дату вручную', callback_data='hwadd_changedate_manual_new'),
                InlineKeyboardButton('❌ Отмена', callback_data='cancel_action'),
            )
            bot.send_message(
                chat_id,
                f'📘 <b>Предмет:</b> {esc(subj_name)}\n\n'
                f'⚠️ Ближайших пар по этому предмету в расписании не нашлось. Выбери дату, к которой прикрепить ДЗ:',
                reply_markup=markup,
                parse_mode='HTML',
            )
        return

    if data.startswith('hwadd_setdate_'):
        if not can_manage_hw(call.from_user):
            return
        _, state = get_user_state(chat_id)
        if not state or 'hw_subject' not in state:
            return
        new_date_str = data.replace('hwadd_setdate_', '')
        try:
            bot.delete_message(chat_id, msg_id)
        except Exception:
            pass
        ask_hw_content(
            chat_id,
            state['hw_group'],
            new_date_str,
            state['hw_subject'],
            upcoming_dates=state.get('hw_upcoming_dates', []),
        )
        return

    if data == 'hwadd_changedate_manual':
        if not can_manage_hw(call.from_user):
            return
        _, state = get_user_state(chat_id)
        if not state or 'hw_subject' not in state:
            return
        group_name = state['hw_group']
        course = find_course_by_group(group_name)
        clear_user_state(chat_id)
        set_user_state(chat_id, 'awaiting_manual_date', {
            'action': 'hw_change_date',
            'course': course,
            'group_name': group_name,
            'hw_subject': state['hw_subject'],
            'hw_upcoming_dates': state.get('hw_upcoming_dates', []),
        })
        bot.edit_message_text(
            f'⌨️ Введи нужную дату для предмета <b>«{esc(state["hw_subject"])}»</b> в формате <code>ДД.ММ.ГГГГ</code> '
            f'(например, <code>{(now_local() + timedelta(days=1)).strftime("%d.%m.%Y")}</code>)\n'
            f'<i>Для отмены напиши «Отмена».</i>',
            chat_id,
            msg_id,
            parse_mode='HTML',
        )
        return

    if data == 'hwadd_changedate_manual_new':
        if not can_manage_hw(call.from_user):
            return
        _, state = get_user_state(chat_id)
        if not state or 'hw_pending_custom_subj' not in state:
            return
        group_name = state['hw_group']
        course = state.get('hw_course') or find_course_by_group(group_name)
        subj_name = state['hw_pending_custom_subj']
        clear_user_state(chat_id)
        set_user_state(chat_id, 'awaiting_manual_date', {
            'action': 'hw_change_date',
            'course': course,
            'group_name': group_name,
            'hw_subject': subj_name,
            'hw_upcoming_dates': [],
        })
        bot.edit_message_text(
            f'⌨️ Введи дату для предмета <b>«{esc(subj_name)}»</b> в формате <code>ДД.ММ.ГГГГ</code>:\n'
            f'<i>Для отмены напиши «Отмена».</i>',
            chat_id,
            msg_id,
            parse_mode='HTML',
        )
        return

    if data.startswith('hwadd_custdate_'):
        if not can_manage_hw(call.from_user):
            return
        _, state = get_user_state(chat_id)
        if not state or 'hw_pending_custom_subj' not in state:
            return
        date_str = data.replace('hwadd_custdate_', '')
        try:
            bot.delete_message(chat_id, msg_id)
        except Exception:
            pass
        ask_hw_content(chat_id, state['hw_group'], date_str, state['hw_pending_custom_subj'])
        return

    if data.startswith('hwadd_date_'):
        if not can_manage_hw(call.from_user):
            return
        course, group_name = get_managed_context(call.from_user)
        if not group_name:
            return
        date_str = data.replace('hwadd_date_', '')
        show_hw_subject_selection(chat_id, msg_id, course, group_name, date_str)
        return

    if data == 'hwadd_datemanual':
        if not can_manage_hw(call.from_user):
            return
        course, group_name = get_managed_context(call.from_user)
        if not group_name:
            return
        clear_user_state(chat_id)
        set_user_state(chat_id, 'awaiting_manual_date', {
            'action': 'hw_add',
            'course': course,
            'group_name': group_name,
        })
        bot.edit_message_text(
            f'⌨️ Введи дату дедлайна в формате <code>ДД.ММ.ГГГГ</code> (например, <code>{(now_local() + timedelta(days=1)).strftime("%d.%m.%Y")}</code>)\n'
            f'<i>Для отмены напиши «Отмена».</i>',
            chat_id,
            msg_id,
            parse_mode='HTML',
        )
        return

    if data == 'hwadd_allsubj':
        if not can_manage_hw(call.from_user):
            return
        _, state = get_user_state(chat_id)
        if not state:
            return
        show_hw_subject_selection(
            chat_id,
            msg_id,
            state['hw_course'],
            state['hw_group'],
            state['hw_date'],
            show_all=True,
        )
        return

    if data == 'hwadd_subjmanual':
        if not can_manage_hw(call.from_user):
            return
        _, state = get_user_state(chat_id)
        if not state:
            return
        clear_user_state(chat_id, keep_payload=True)
        set_user_state(chat_id, 'awaiting_hw_manual_subj', state)
        bot.edit_message_text(
            f'✍️ Введи название предмета для ДЗ на <b>{state["hw_date"]}</b>:\n'
            f'<i>Для отмены напиши «Отмена».</i>',
            chat_id,
            msg_id,
            parse_mode='HTML',
        )
        return

    if data.startswith('hwadd_subj_'):
        if not can_manage_hw(call.from_user):
            return
        _, state = get_user_state(chat_id)
        if not state:
            return
        idx = int(data.replace('hwadd_subj_', ''))
        subjects = state.get('hw_subjects', [])
        if idx >= len(subjects):
            return
        subj_name = subjects[idx]
        try:
            bot.delete_message(chat_id, msg_id)
        except Exception:
            pass
        ask_hw_content(chat_id, state['hw_group'], state['hw_date'], subj_name)
        return

    if data.startswith('hw_showdate_'):
        user = get_user_data(user_id)
        if not user:
            return
        date_str = data.replace('hw_showdate_', '')
        send_homeworks_for_date_view(chat_id, call.from_user, user[1], date_str)
        return

    if data == 'hw_view_week':
        user = get_user_data(user_id)
        if not user:
            return
        send_homeworks_for_period(chat_id, call.from_user, user[1], days_limit=7)
        return

    if data == 'hw_view_all':
        user = get_user_data(user_id)
        if not user:
            return
        send_homeworks_for_period(chat_id, call.from_user, user[1], days_limit=None)
        return

    if data == 'hw_view_manual':
        user = get_user_data(user_id)
        if not user:
            return
        clear_user_state(chat_id)
        set_user_state(chat_id, 'awaiting_manual_date', {
            'action': 'hw_view',
            'course': user[0],
            'group_name': user[1],
        })
        bot.send_message(
            chat_id,
            f'🔍 Введи дату для просмотра ДЗ в формате <code>ДД.ММ.ГГГГ</code> (например, <code>{(now_local() + timedelta(days=1)).strftime("%d.%m.%Y")}</code>)\n'
            f'<i>Для отмены напиши «Отмена».</i>',
            parse_mode='HTML',
        )
        return

    if data.startswith('hw_edit_'):
        if not can_manage_hw(call.from_user):
            return
        hw_id = int(data.replace('hw_edit_', ''))

        with db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                'SELECT target_date, subject_name, hw_text, group_name, added_by FROM homeworks '
                'WHERE id = ?',
                (hw_id,),
            )
            hw_row_raw = cursor.fetchone()

        if not hw_row_raw:
            bot.send_message(chat_id, '⛔ Запись не найдена.')
            return
        hw_date, hw_subj, old_text, group_name, hw_added_by = hw_row_raw
        if not can_edit_specific_hw(call.from_user, group_name, hw_added_by):
            bot.send_message(chat_id, '⛔ Редактировать ДЗ может только его автор, староста группы или администратор.')
            return

        clear_user_state(chat_id)
        set_user_state(chat_id, 'awaiting_hw_edit', {'hw_id': hw_id, 'group_name': group_name})
        bot.send_message(
            chat_id,
            f'✏️ <b>РЕДАКТИРОВАНИЕ ТЕКСТА ДЗ</b>\n'
            f'<blockquote>📅 <b>Дедлайн:</b> {hw_date}\n'
            f'📘 <b>Предмет:</b> {esc(hw_subj)}</blockquote>\n\n'
            f'<b>Текущий текст:</b>\n<code>{esc(old_text or "(пусто)")}</code>\n\n'
            f'Отправь новый текст задания (все прикреплённые фото и документы сохранятся).\n'
            f'<i>Для отмены напиши «Отмена».</i>',
            parse_mode='HTML',
        )
        return

    if data.startswith('hw_del_'):
        if not can_manage_hw(call.from_user):
            return
        hw_id = int(data.replace('hw_del_', ''))

        with db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                'SELECT target_date, subject_name, hw_text, queue_id, group_name, added_by FROM homeworks '
                'WHERE id = ?',
                (hw_id,),
            )
            hw_row_raw = cursor.fetchone()
            if not hw_row_raw:
                bot.send_message(chat_id, '⛔ Запись не найдена.')
                return
            hw_date, hw_subj, hw_text, linked_q_id, group_name, hw_added_by = hw_row_raw
            if not can_edit_specific_hw(call.from_user, group_name, hw_added_by):
                bot.send_message(chat_id, '⛔ Удалить ДЗ может только его автор, староста группы или администратор.')
                return

            if linked_q_id:
                cursor.execute('DELETE FROM queue_slots WHERE queue_id = ?', (linked_q_id,))
                cursor.execute('DELETE FROM queues WHERE id = ?', (linked_q_id,))

            cursor.execute(
                'DELETE FROM homeworks WHERE id = ? AND group_name = ?',
                (hw_id, group_name),
            )
            cursor.execute(
                'SELECT user_id FROM users WHERE group_name = ? AND notify_updates = 1',
                (group_name,),
            )
            group_users = cursor.fetchall()

        try:
            bot.delete_message(chat_id, msg_id)
        except Exception:
            pass
        bot.send_message(chat_id, '🗑 Домашнее задание удалено.')

        notify_text = (
            f'🗑 <b>ДОМАШНЕЕ ЗАДАНИЕ УДАЛЕНО</b>\n'
            f'<blockquote>🎓 <b>Группа:</b> {esc(group_name)}\n'
            f'📅 <b>Дедлайн был:</b> {hw_date}</blockquote>\n\n'
            f'📘 <b>{esc(hw_subj)}</b>\n'
            f'<s>{esc(hw_text) if hw_text else "(вложение)"}</s>'
        )
        for (u_id,) in group_users:
            if u_id != user_id:
                enqueue_message(u_id, notify_text, parse_mode='HTML')
        return


# --- ЕДИНЫЙ РОУТЕР МАШИНЫ СОСТОЯНИЙ (FSM НА БАЗЕ SQLITE user_states) ---
STATE_HANDLERS = {
    'awaiting_hw_content': process_hw_content,
    'awaiting_hw_slots_count': process_hw_queue_count_input,
    'awaiting_hw_manual_subj': process_hw_manual_subject,
    'awaiting_hw_edit': process_hw_edit_text,
    'awaiting_repl_subject': process_replacement_subject,
    'awaiting_manual_date': process_manual_date_for_action,
    'awaiting_broadcast': process_broadcast_content,
    'awaiting_teacher_query': process_teacher_search,
    'awaiting_docx_date': process_docx_custom_start_date,
}


@bot.message_handler(
    content_types=['text', 'photo', 'document'],
    func=lambda m: bool(m.from_user and get_user_state(m.from_user.id)[0] in STATE_HANDLERS),
)
def fsm_message_router(message):
    sync_username(message.from_user)
    state_name, payload = get_user_state(message.from_user.id)
    handler = STATE_HANDLERS.get(state_name)
    if handler:
        handler(message, payload)


# --- ПРОВЕРКА ЗАВИСШИХ СОСТОЯНИЙ ПРИ ЗАПУСКЕ БОТА ---
def restore_interrupted_dialogs():
    now_ts = time.time()
    with db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('SELECT user_id, state, payload_json, updated_at FROM user_states WHERE state != ""')
        rows = cursor.fetchall()

    for u_id, state_name, payload_json, updated_at in rows:
        if now_ts - float(updated_at or 0) > STATE_TTL_SECONDS or state_name not in STATE_HANDLERS:
            clear_user_state(u_id)
            continue
        if state_name == 'awaiting_docx_date':
            try:
                payload = json.loads(payload_json) if payload_json else {}
            except Exception:
                payload = {}
            if not get_pending_docx_import(payload.get('import_token', '')):
                clear_user_state(u_id)


# --- АВТОМАТИЧЕСКАЯ РАССЫЛКА (07:00 И 18:00) ЧЕРЕЗ ФОНОВУЮ ОЧЕРЕДЬ ---
def _run_scheduled_mailing(day_offset, greeting_line, day_label_genitive):
    target_dt = now_local() + timedelta(days=day_offset)
    target_str = target_dt.strftime('%d.%m.%Y')
    with db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT user_id, course, group_name, faculty FROM users "
            "WHERE subscribed = 1 AND group_name IS NOT NULL AND group_name != '' AND course > 0"
        )
        users = cursor.fetchall()

    group_cache = {}
    for u_id, course, group_name, fac_code in users:
        cache_key = (fac_code, course, group_name)
        if cache_key not in group_cache:
            text = (
                f'{greeting_line}\n\n'
                + get_schedule_text(target_dt, course, group_name, faculty_code=fac_code)
            )
            kb = get_pagination_keyboard(day_offset, group_name)
            attachments = []
            for _, hw_subj, hw_text, hw_photo, hw_doc_id, hw_doc_name in get_homeworks_for_date(group_name, target_str):
                photo_ids = parse_photo_ids(hw_photo)
                if photo_ids or hw_doc_id:
                    cap = (
                        f'📎 <b>Вложение к ДЗ на {day_label_genitive} ({target_str})</b>\n'
                        f'📘 <b>{esc(hw_subj)}</b>\n'
                        f'📝 {esc(hw_text) if hw_text else ""}'
                    ).strip()
                    attachments.append((photo_ids, cap, hw_doc_id, hw_doc_name))
            group_cache[cache_key] = (text, kb, attachments)

        text, kb, attachments = group_cache[cache_key]
        enqueue_message(u_id, text, reply_markup=kb, parse_mode='HTML')
        for photo_ids, cap, hw_doc_id, hw_doc_name in attachments:
            enqueue_hw_media(u_id, photo_ids, cap, doc_file_id=hw_doc_id, doc_file_name=hw_doc_name)


def morning_mailing():
    _run_scheduled_mailing(
        0,
        '☀️ <b>Доброе утро!</b> Актуальное расписание пар и ДЗ на сегодня:',
        'сегодня',
    )


def evening_mailing():
    _run_scheduled_mailing(
        1,
        '🌙 <b>Добрый вечер!</b> Напоминание о расписании пар и ДЗ на завтра:',
        'завтра',
    )


scheduler = BackgroundScheduler(timezone='Asia/Yekaterinburg')
scheduler.add_job(morning_mailing, 'cron', day_of_week='mon-fri', hour=7, minute=0)
scheduler.add_job(evening_mailing, 'cron', day_of_week='sun,mon,tue,wed,thu', hour=18, minute=0)
scheduler.add_job(send_db_backup, 'cron', hour=3, minute=0)
scheduler.add_job(cleanup_stale_states, 'interval', minutes=30)
scheduler.add_job(cleanup_old_db_records, 'cron', day_of_week='sun', hour=4, minute=0)


if __name__ == '__main__':
    if not TOKEN or TOKEN == 'YOUR_BOT_TOKEN_HERE':
        raise RuntimeError('Не задан токен бота! Укажи переменную окружения BOT_TOKEN или впиши токен в TOKEN.')
    init_db()
    restore_interrupted_dialogs()
    scheduler.start()
    try:
        # Обычным пользователям в личке показываем только /start
        bot.set_my_commands([
            BotCommand('start', 'Перезапустить бота / главное меню'),
        ])
        # В групповых чатах полностью очищаем список команд по "/", чтобы не мозолили глаза
        bot.delete_my_commands(scope=BotCommandScopeAllGroupChats())
        # Админу устанавливаем персональное меню команд только в его личном чате
        with db_connection() as conn:
            c_adm = conn.cursor()
            c_adm.execute(
                'SELECT user_id FROM users WHERE LOWER(username) = ?',
                (ADMIN_USERNAME.lower(),),
            )
            for (adm_id,) in c_adm.fetchall():
                sync_admin_personal_commands(adm_id)
    except Exception as e:
        print(f'Не удалось установить меню команд: {e}')
    print('Бот СИ УУНиТ успешно запущен и работает!')
    while True:
        try:
            bot.infinity_polling(
                timeout=30,
                long_polling_timeout=25,
                allowed_updates=['message', 'callback_query', 'inline_query'],
            )
        except Exception as e:
            print(f'Сбой соединения polling, перезапуск через 5 сек: {e}')
            time.sleep(5)
