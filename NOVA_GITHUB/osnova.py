"""
Бот-помощник Nova для MAX (aiomax).

Команды:
  /game      - собрать презентацию из вопросов и ответов (/game 5 - 5 пар)
  /random    - разбить людей на группы (/reshuffle - перемешать тех же людей)
  /ai        - режим диалога с ИИ «Алиса» (/ai вопрос - разовый ответ)
  /help      - список команд
  /cancel    - отменить текущее действие

Чтобы добавить новую функцию: напишите async-функцию (message, args)
и повесьте на неё @command("/имя", "Описание"). Для диалога из нескольких
шагов используйте set_state(...) - см. /random.

Зависимости: pip install aiomax python-dotenv python-pptx
Токен: файл .env рядом с bot-файлом, строка BOT_TOKEN=...
"""
import asyncio
import logging
import os
import random
import re
from pathlib import Path
from urllib.parse import quote

import aiohttp
import aiomax
from dotenv import load_dotenv
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Pt

# ============ НАСТРОЙКИ ============
BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / "TOKEN.env")

TOKEN = os.getenv("BOT_TOKEN")
if not TOKEN or set(TOKEN) == {"*"}:
    raise SystemExit("BOT_TOKEN не найден. Впишите токен в файл  TOKEN.env рядом с ботом: BOT_TOKEN=ваш_токен")

TEMPLATE_PATH = BASE_DIR / "template.pptx"
OUTPUT_DIR = BASE_DIR / "output"
TOTAL_PAIRS = 9                      # сколько пар вопрос-ответ по умолчанию

QUESTION_PLACEHOLDER = "ТЕКСТ ВОПРОСА"
ANSWER_PLACEHOLDER = "ТЕКСТ ОТВЕТА"

AI_URL = "https://text.pollinations.ai/"
AI_PROMPT = "Ты Алиса, помощник учителя. Отвечай коротко и по-русски. Вопрос: "
MESSAGE_LIMIT = 4000                 # максимум символов в одном сообщении MAX

bot = aiomax.Bot(TOKEN, default_format="markdown", use_certificate=True)

# ============ ЯДРО: КОМАНДЫ И ДИАЛОГИ ============
commands = {}   # "/game" -> (функция, описание); пустое описание = скрыто в /help
states = {}     # user_id -> (обработчик следующего сообщения, данные)


def command(name, description="", aliases=()):
    """Регистрирует команду: @command("/game", "Описание")."""
    def decorator(func):
        commands[name.lower()] = (func, description)
        for alias in aliases:
            commands[alias.lower()] = (func, "")
        return func
    return decorator


def set_state(user_id, handler, data=None):
    """Следующие сообщения пользователя пойдут в handler(message, data)."""
    states[user_id] = (handler, data if data is not None else {})


def clear_state(user_id):
    states.pop(user_id, None)


async def send_long(message, text):
    """Отправляет длинный текст частями (лимит сообщения MAX)."""
    text = text.strip() or "(пустой ответ)"
    for i in range(0, len(text), MESSAGE_LIMIT):
        await message.send(text[i:i + MESSAGE_LIMIT])


async def send_file(message, path, text):
    """Загружает файл и отправляет в чат (с повторами, пока MAX обрабатывает файл)."""
    with open(path, "rb") as f:
        attachment = await bot.upload_file(f)

    last_error = None
    for _ in range(5):
        try:
            await message.send(text, attachments=[attachment])
            return
        except Exception as e:
            last_error = e
            await asyncio.sleep(1.5)
    raise last_error


@bot.on_message()
async def dispatch(message: aiomax.Message):
    user_id = message.sender.user_id
    text = (message.content or "").strip()
    parts = text.split(maxsplit=1)
    cmd = parts[0].lower() if parts else ""
    args = parts[1].strip() if len(parts) > 1 else ""

    try:
        if cmd in commands:
            clear_state(user_id)  # новая команда прерывает прошлый диалог
            await commands[cmd][0](message, args)
        elif user_id in states:
            handler, data = states[user_id]
            await handler(message, data)
        elif cmd.startswith("/"):
            await message.send("Не знаю такой команды. Список команд: /help")
        else:
            await message.send("Не понял. Список команд: /help")
    except Exception as e:
        logging.exception("Ошибка в обработчике")
        clear_state(user_id)
        await message.send(f"Произошла ошибка: {e}")


@bot.on_bot_start()
async def on_start(payload: aiomax.BotStartPayload):
    await payload.send("Здравствуйте! Я - бот-помощник от команды Nova. Список команд: /help")


# ============ ОБЩИЕ КОМАНДЫ ============
@command("/help", "Список команд", aliases=("/start",))
async def help_cmd(message, args):
    lines = [f"{name} - {desc}" for name, (_, desc) in commands.items() if desc]
    await message.send("Доступные команды:\n" + "\n".join(lines))


@command("/cancel", "Отменить текущее действие")
async def cancel(message, args):
    clear_state(message.sender.user_id)
    await message.send("Действие отменено.")


# ============ ФУНКЦИЯ: /ai (Алиса) ============
async def ask_ai(question):
    """Спрашивает ИИ и возвращает текст ответа."""
    url = AI_URL + quote(AI_PROMPT + question[:1500], safe="")
    timeout = aiohttp.ClientTimeout(total=60)
    async with aiohttp.ClientSession(timeout=timeout) as session:
        async with session.get(url) as resp:
            resp.raise_for_status()
            return (await resp.text()).strip()


async def answer_ai(message, question):
    try:
        answer = await ask_ai(question)
    except Exception:
        logging.exception("Ошибка запроса к ИИ")
        await message.send("Алиса сейчас не отвечает, попробуйте ещё раз чуть позже.")
        return
    await send_long(message, answer)


@command("/ai", "Спросить ИИ: /ai вопрос или режим диалога")
async def start_ai(message, args):
    if args:                                    # разовый вопрос
        await answer_ai(message, args)
        return
    set_state(message.sender.user_id, ai_step)
    await message.send("Привет! Я Алиса. Задайте вопрос. Выйти из диалога - /cancel")


async def ai_step(message, data):
    question = (message.content or "").strip()
    if question:
        await answer_ai(message, question)


# ============ ФУНКЦИЯ: /random (разбивка на группы) ============
last_random = {}   # user_id -> (люди, группы) для /reshuffle


def parse_list(text):
    """
    Превращает ввод в список. Понимает:
      '1.Петр Иванов 2.Иван Жук', 'Петр Иванов, Иван Жук',
      по одному в строке, а также простые слова через пробел ('Красные Синие').
    """
    text = text.strip()
    if re.search(r"[,;\n]", text):
        parts = [re.sub(r"^\s*\d+\s*[.)]\s*", "", p) for p in re.split(r"[,;\n]+", text)]
    elif re.search(r"(?:^|\s)\d+\s*[.)]", text):
        parts = re.split(r"(?:^|\s)\d+\s*[.)]\s*", text)
    else:
        parts = text.split()
    return [p.strip() for p in parts if p.strip()]


def make_groups(names, people):
    """Раздаёт людей по группам случайно, разница в размерах не больше 1."""
    shuffled = random.sample(people, len(people))
    groups = [[] for _ in names]
    for i, person in enumerate(shuffled):
        groups[i % len(names)].append(person)
    return "\n".join(
        f"**{name}** ({len(members)}): {', '.join(members)}"
        for name, members in zip(names, groups)
    )


@command("/random", "Разбить людей на группы")
async def start_random(message, args):
    set_state(message.sender.user_id, random_people_step)
    await message.send(
        "Напишите список людей - с номерами, через запятую или по одному в строке:\n"
        "1.Петр Иванов 2.Иван Жук 3.Маша Петрова"
    )


async def random_people_step(message, data):
    people = parse_list(message.content or "")
    if len(people) < 2:
        await message.send("Нужно минимум 2 человека. Попробуйте ещё раз (или /cancel).")
        return
    set_state(message.sender.user_id, random_groups_step, {"people": people})
    await message.send(
        f"Нашёл {len(people)} чел. Теперь напишите названия групп, например: 1.Красные 2.Синие\n"
        "Или просто число - тогда группы будут «Группа 1», «Группа 2»..."
    )


async def random_groups_step(message, data):
    text = (message.content or "").strip()
    people = data["people"]

    if text.isdigit():
        groups = [f"Группа {i}" for i in range(1, int(text) + 1)]
    else:
        groups = parse_list(text)

    if len(groups) < 2:
        await message.send("Нужно минимум 2 группы. Попробуйте ещё раз (или /cancel).")
        return
    if len(groups) > len(people):
        await message.send(f"Групп ({len(groups)}) больше, чем людей ({len(people)}). Уменьшите число групп.")
        return

    user_id = message.sender.user_id
    clear_state(user_id)
    last_random[user_id] = (people, groups)
    await message.send(make_groups(groups, people))
    await message.send("Другой состав групп - /reshuffle, новый список - /random")


@command("/reshuffle", "Перемешать тех же людей по тем же группам")
async def reshuffle(message, args):
    saved = last_random.get(message.sender.user_id)
    if not saved:
        await message.send("Сначала составьте группы: /random")
        return
    people, groups = saved
    await message.send(make_groups(groups, people))


# ============ PPTX ============
def iter_shapes(shapes):
    """Обходит все фигуры, включая сгруппированные."""
    for shape in shapes:
        try:
            is_group = shape.shape_type == MSO_SHAPE_TYPE.GROUP
        except NotImplementedError:
            is_group = False
        if is_group:
            yield from iter_shapes(shape.shapes)
        else:
            yield shape


def find_shape(slide, placeholder):
    """Ищет фигуру, в тексте которой есть заготовка."""
    for shape in iter_shapes(slide.shapes):
        if shape.has_text_frame and placeholder in shape.text_frame.text.upper():
            return shape
    return None


def replace_placeholder(slide, placeholder, new_text):
    """Заменяет текст заготовки, сохраняя форматирование."""
    shape = find_shape(slide, placeholder)
    if shape is None:
        return False
    paragraph = shape.text_frame.paragraphs[0]
    runs = paragraph.runs
    if runs:
        runs[0].text = new_text
        for extra in runs[1:]:
            extra._r.getparent().remove(extra._r)
    else:
        paragraph.text = new_text
    for extra_p in shape.text_frame.paragraphs[1:]:
        extra_p._p.getparent().remove(extra_p._p)
    return True


def add_fallback_textbox(prs, slide, text, top_ratio):
    w, h = prs.slide_width, prs.slide_height
    box_w, box_h = int(w * 0.8), int(h * 0.25)
    box = slide.shapes.add_textbox((w - box_w) // 2, int(h * top_ratio), box_w, box_h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(28)


def delete_slide(prs, index):
    slide_ids = prs.slides._sldIdLst
    slide_id = slide_ids[index]
    prs.part.drop_rel(slide_id.rId)
    slide_ids.remove(slide_id)


def count_template_pairs():
    """Сколько пар вопрос-ответ помещается в шаблон (слайд 1 - титульный)."""
    return max(len(Presentation(TEMPLATE_PATH).slides) - 1, 0)


def build_presentation(questions, answers, user_id):
    """Подставляет пары вопрос-ответ в слайды 2, 3, 4... и сохраняет новый файл."""
    prs = Presentation(TEMPLATE_PATH)
    slides = prs.slides

    for i, (question, answer) in enumerate(zip(questions, answers)):
        slide = slides[i + 1]  # пара 1 -> слайд 2
        if not replace_placeholder(slide, QUESTION_PLACEHOLDER, question):
            logging.warning("Слайд %d: не найдена заготовка вопроса", i + 2)
            add_fallback_textbox(prs, slide, question, 0.35)
        if not replace_placeholder(slide, ANSWER_PLACEHOLDER, answer):
            logging.warning("Слайд %d: не найдена заготовка ответа", i + 2)
            add_fallback_textbox(prs, slide, answer, 0.70)

    # Убираем неиспользованные слайды с пустыми заготовками (если пар меньше, чем слайдов)
    for idx in reversed(range(len(questions) + 1, len(slides))):
        s = slides[idx]
        if find_shape(s, QUESTION_PLACEHOLDER) or find_shape(s, ANSWER_PLACEHOLDER):
            delete_slide(prs, idx)

    OUTPUT_DIR.mkdir(exist_ok=True)
    output = OUTPUT_DIR / f"presentation_{user_id}.pptx"
    prs.save(output)
    return output


# ============ ФУНКЦИЯ: /game ============
@command("/game", "Собрать презентацию из вопросов и ответов (/game 5 - пять пар)")
async def start_game(message, args):
    if not TEMPLATE_PATH.exists():
        await message.send(f"Не найден шаблон {TEMPLATE_PATH.name} рядом с ботом.")
        return

    capacity = count_template_pairs()
    total = int(args) if args.isdigit() else TOTAL_PAIRS
    total = max(1, min(total, capacity))

    set_state(message.sender.user_id, game_step, {"questions": [], "answers": [], "total": total})
    await message.send(f"Введите вопрос 1 из {total}:")


async def game_step(message, data):
    text = (message.content or "").strip()
    if not text:
        await message.send("Нужен текст, попробуйте ещё раз.")
        return

    questions, answers, total = data["questions"], data["answers"], data["total"]

    if len(questions) == len(answers):      # ждали вопрос
        questions.append(text)
        await message.send(f"Введите ответ {len(questions)} из {total}:")
        return

    answers.append(text)                    # ждали ответ
    if len(answers) < total:
        await message.send(f"Введите вопрос {len(answers) + 1} из {total}:")
        return

    user_id = message.sender.user_id
    clear_state(user_id)
    await message.send("Собираю презентацию...")

    path = await asyncio.to_thread(build_presentation, questions, answers, user_id)
    logging.info("Презентация сохранена: %s", path)
    await send_file(message, path, "Готово! Вот ваша презентация:")


# ============ ЗАПУСК ============
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    bot.run()
