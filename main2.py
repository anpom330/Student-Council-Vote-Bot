import pandas as pd
import json
from datetime import datetime
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, CallbackQueryHandler, ContextTypes, PollAnswerHandler
from telegram.ext import MessageHandler, filters
import io
from docx import Document
from docx4 import create_hse_docx

# Список разрешенных пользователей (user_id)
ALLOWED_USERS = [527088298, 881860095]  # Замените на реальные user_id

user_poll_data = {}

# Файл для хранения данных
DATA_FILE = "poll_data3.json"


async def start_custom_poll(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()  # Убирает "загрузку"

    # Проверка доступа
    if not await check_access(update):
        return

    user_id = query.from_user.id

    # Инициализация данных для текущего пользователя
    user_poll_data[user_id] = {"stage": "WAITING_FOR_QUESTION"}
    await query.message.reply_text("Введите вопрос для опроса:")




async def handle_custom_poll2(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    message_text = update.message.text.strip()

    # Проверка и инициализация данных пользователя
    if user_id not in user_poll_data:
        print(f"[DEBUG] Получен вопрос от пользователя {user_id}: {message_text}")
        user_poll_data[user_id] = {
            "question": message_text,  # Сохраняем вопрос
            "options": [],            # Пустой список для вариантов
            "state": "options"        # Устанавливаем состояние "options"
        }
        print(f"[DEBUG] Инициализация данных опроса для пользователя {user_id}: {user_poll_data[user_id]}")
        await update.message.reply_text(
            "Введите варианты ответа по одному. Когда закончите, напишите 'Готово', чтобы завершить опрос."
        )
        return

    # Проверка целостности структуры данных пользователя
    if "state" not in user_poll_data[user_id] or "question" not in user_poll_data[user_id]:
        print(f"[WARNING] Нарушена структура данных для пользователя {user_id}. Переинициализация.")
        user_poll_data[user_id] = {
            "question": message_text,
            "options": [],
            "state": "options"
        }

    # Состояние пользователя
    user_state = user_poll_data[user_id]["state"]

    print(f"[DEBUG] Получен ввод от пользователя {user_id}: '{message_text}'")
    print(f"[DEBUG] Текущее состояние пользователя {user_id}: {user_state}")

    # Обработка ввода при добавлении вариантов
    if user_state == "options":
        if message_text.lower() == "готово":
            # Проверка наличия вариантов
            if len(user_poll_data[user_id]["options"]) < 2:
                await update.message.reply_text("Опрос должен содержать как минимум два варианта ответа.")
                print(f"[ERROR] Пользователь {user_id} указал недостаточно вариантов.")
                return

            # Завершение и отправка опроса
            user_poll_data[user_id]["state"] = "completed"
            question = user_poll_data[user_id]["question"]
            options = user_poll_data[user_id]["options"]

            print(f"[DEBUG] Завершение опроса пользователя {user_id}. Вопрос: {question}, Варианты: {options}")
            message = await update.message.reply_poll(
                question=question,
                options=options,
                is_anonymous=False
            )
            print(f"[DEBUG] Опрос отправлен. ID: {message.poll.id}")

            # Сохранение данных опроса
            poll_data[message.poll.id] = {
                "question": question,
                "options": options,
                "votes": {option: [] for option in options}
            }
            save_poll_data()
            print(f"[DEBUG] Опрос сохранен с ID {message.poll.id}: {poll_data[message.poll.id]}")

            await update.message.reply_text("Опрос успешно создан!")
            del user_poll_data[user_id]
            print(f"[DEBUG] Данные пользователя {user_id} удалены после создания опроса.")
            return

        # Добавление варианта
        if message_text not in user_poll_data[user_id]["options"]:
            user_poll_data[user_id]["options"].append(message_text)
            print(f"[DEBUG] Пользователь {user_id} добавил вариант: {message_text}")
            await update.message.reply_text(f"Вариант '{message_text}' добавлен. Введите следующий вариант или 'Готово' для завершения.")
        else:
            print(f"[WARNING] Пользователь {user_id} попытался добавить дублирующий вариант: {message_text}")
            await update.message.reply_text(f"Вариант '{message_text}' уже добавлен. Введите другой вариант.")
        return

    # Обработка повторного ввода после завершения
    if user_state == "completed":
        print(f"[WARNING] Пользователь {user_id} попытался создать новый опрос, не начав заново.")
        await update.message.reply_text("Опрос уже завершен. Для создания нового начните с ввода нового вопроса.")
        return



async def handle_custom_poll(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    message_text = update.message.text.strip()

    # Проверка и инициализация данных пользователя
    if user_id not in user_poll_data:
        print(f"[DEBUG] Получен вопрос от пользователя {user_id}: {message_text}")
        user_poll_data[user_id] = {
            "question": message_text,  # Сохраняем текст как вопрос
            "options": [],            # Создаём пустой список для вариантов
            "state": "options"        # Устанавливаем состояние "options"
        }
        print(f"[DEBUG] Инициализация данных опроса для пользователя {user_id}: {user_poll_data[user_id]}")
        await update.message.reply_text(
            "Введите варианты ответа по одному. Когда закончите, напишите 'Готово', чтобы завершить опрос."
        )
        return

    # Проверка целостности структуры данных пользователя
    if "state" not in user_poll_data[user_id] or "question" not in user_poll_data[user_id]:
        print(f"[WARNING] Нарушена структура данных для пользователя {user_id}. Переинициализация.")
        user_poll_data[user_id] = {
            "question": message_text,
            "options": [],
            "state": "options"
        }
        await update.message.reply_text(
            "Введите варианты ответа по одному. Когда закончите, напишите 'Готово', чтобы завершить опрос."
        )
        return

    # Состояние пользователя
    user_state = user_poll_data[user_id]["state"]

    print(f"[DEBUG] Получен ввод от пользователя {user_id}: '{message_text}'")
    print(f"[DEBUG] Текущее состояние пользователя {user_id}: {user_state}")

    # Обработка ввода при добавлении вариантов
    if user_state == "options":
        if message_text.lower() == "готово":
            # Проверка наличия вариантов
            if len(user_poll_data[user_id]["options"]) < 2:
                await update.message.reply_text("Опрос должен содержать как минимум два варианта ответа.")
                print(f"[ERROR] Пользователь {user_id} указал недостаточно вариантов.")
                return

            # Завершение и отправка опроса
            user_poll_data[user_id]["state"] = "completed"
            question = user_poll_data[user_id]["question"]
            options = user_poll_data[user_id]["options"]

            print(f"[DEBUG] Завершение опроса пользователя {user_id}. Вопрос: {question}, Варианты: {options}")
            message = await update.message.reply_poll(
                question=question,
                options=options,
                is_anonymous=False
            )
            print(f"[DEBUG] Опрос отправлен. ID: {message.poll.id}")

            # Сохранение данных опроса
            poll_data[message.poll.id] = {
                "question": question,
                "options": options,
                "votes": {option: [] for option in options}
            }
            save_poll_data()
            print(f"[DEBUG] Опрос сохранен с ID {message.poll.id}: {poll_data[message.poll.id]}")

            await update.message.reply_text("Опрос успешно создан!")
            del user_poll_data[user_id]
            print(f"[DEBUG] Данные пользователя {user_id} удалены после создания опроса.")
            return

        # Добавление варианта
        if message_text not in user_poll_data[user_id]["options"]:
            user_poll_data[user_id]["options"].append(message_text)
            print(f"[DEBUG] Пользователь {user_id} добавил вариант: {message_text}")
            await update.message.reply_text(f"Вариант '{message_text}' добавлен. Введите следующий вариант или 'Готово' для завершения.")
        else:
            print(f"[WARNING] Пользователь {user_id} попытался добавить дублирующий вариант: {message_text}")
            await update.message.reply_text(f"Вариант '{message_text}' уже добавлен. Введите другой вариант.")
        return

    # Обработка повторного ввода после завершения
    if user_state == "completed":
        print(f"[WARNING] Пользователь {user_id} попытался создать новый опрос, не начав заново.")
        await update.message.reply_text("Опрос уже завершен. Для создания нового начните с ввода нового вопроса.")
        return


# Загрузка данных из таблицы голосов
def load_poll_data():
    try:
        with open(DATA_FILE, "r") as file:
            return json.load(file)
    except FileNotFoundError:
        return {}

poll_data = load_poll_data()
def save_poll_data():
    with open(DATA_FILE, "w") as file:
        json.dump(poll_data, file, indent=4)

# Инициализация данных опросов


async def check_access(update: Update) -> bool:
    # Извлечение user_id в зависимости от типа объекта
    if isinstance(update, Update) and update.message:
        user_id = update.message.from_user.id
    elif isinstance(update, Update) and update.callback_query:
        user_id = update.callback_query.from_user.id
    else:
        return False  # Если объект не соответствует ни одному типу

    if user_id not in ALLOWED_USERS:
        if update.message:
            await update.message.reply_text("У вас нет доступа к этому боту.")
        elif update.callback_query:
            await update.callback_query.answer("У вас нет доступа к этому боту.", show_alert=True)
        return False
    return True

# Команда /start
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await check_access(update):
        return

    keyboard = [
        [InlineKeyboardButton("Создать тестовый опрос", callback_data="create_test_poll")],
        [InlineKeyboardButton("Создать опрос", callback_data="create_poll")],
        [InlineKeyboardButton("Посмотреть результаты", callback_data="view_results")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text("Привет! Выберите действие:", reply_markup=reply_markup)



async def create_test_poll(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await check_access(update):
        return

    query = update.callback_query
    await query.answer()

    question = "Какой ваш любимый цвет?"
    options = ["Красный", "Синий", "Зеленый", "Воздержаться", "Тык"]

    message = await query.message.reply_poll(
        question=question,
        options=options,
        is_anonymous=False
    )

    poll_data[message.poll.id] = {
        "question": question,
        "options": options,
        "votes": {option: [] for option in options}
    }
    save_poll_data()

    await query.message.reply_text(f"Тестовый опрос создан. ID: {message.poll.id}")


# Обработка голосов в опросе
async def handle_poll_answer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    answer = update.poll_answer
    poll_id = answer.poll_id
    user = update.effective_user

    if poll_id in poll_data:
        option_id = answer.option_ids[0]
        option_text = poll_data[poll_id]["options"][option_id]

        for voters in poll_data[poll_id]["votes"].values():
            if user.username in voters:
                voters.remove(user.username)

        poll_data[poll_id]["votes"][option_text].append(user.username)
        save_poll_data()


async def view_results(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()  # Завершает callback-запрос, чтобы убрать "загрузку"

    # Проверка доступа
    if not await check_access(update):
        return

    # Получение последних 5 опросов
    last_polls = list(poll_data.keys())[-10:]
    if not last_polls:
        await query.message.reply_text("Нет созданных опросов.")
        return

    # Генерация кнопок для выбора опросов
    keyboard = [
        [InlineKeyboardButton(f"{poll_id} - {poll_data[poll_id]['question']}",
                              callback_data=f"view_poll_{poll_id}")]
        for poll_id in last_polls if poll_id in poll_data  # Проверка, что данные о вопросе есть
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)

    # Редактирование сообщения для отображения кнопок
    await query.message.edit_text('Выберите опрос для просмотра результатов:', reply_markup=reply_markup)


# Показать список проголосовавших




async def show_voters(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    poll_id = query.data.split("_")[-1]

    # Проверка на существование опроса
    if poll_id not in poll_data:
        await query.message.reply_text("Опрос не найден.")
        return

    # Загрузка данных из таблицы
    try:
        df = pd.read_excel("voters_data.xlsx")
    except Exception as e:
        await query.message.reply_text(f"Ошибка при чтении файла: {e}")
        return

    # Проверка наличия необходимых столбцов
    required_columns = {"name", "department", "username", "weight"}
    if not required_columns.issubset(df.columns):
        await query.message.reply_text(f"В таблице отсутствуют необходимые столбцы: {required_columns - set(df.columns)}")
        return

    # Подсчет общего веса голосов всех людей в таблице
    total_weight = df["weight"].sum()

    # Подсчет количества человек в студсовете (непустые значения weight)
    council_members_count = df["weight"].notna().sum()

    # Получение данных по голосованию
    votes = poll_data[poll_id]["votes"]
    options = poll_data[poll_id]["options"]
    results = {}
    total_voted_weight = 0  # Суммарный вес голосов среди проголосовавших
    total_voters_count = 0  # Количество проголосовавших

    # Обработка каждого варианта ответа
    for option in options[:-1]:  # Игнорируем последний вариант ответа
        voters = votes.get(option, [])
        total_option_weight = 0
        voter_details = []

        for username in voters:
            user_row = df[df["username"] == username]

            if not user_row.empty:
                weight = user_row.iloc[0]["weight"]
                name = user_row.iloc[0]["name"]
                department = user_row.iloc[0]["department"]
                total_option_weight += weight
                voter_details.append(f"{name}")
                total_voted_weight += weight
                total_voters_count += 1
            else:
                await query.message.reply_text(f"Внимание: Пользователь @{username} отсутствует в таблице и не был учтен.")

        results[option] = {
            "total_weight": total_option_weight,
            "voters": voter_details
        }

    # Формирование сообщения с результатами
    message = (
        f"Количество человек в студсовете: {council_members_count}\n"
        f"Количество голосов среди всех членов: {total_weight}\n"
        f"Количество голосов среди проголосовавших: {total_voted_weight}\n"
        f"Количество проголосовавших: {total_voters_count}\n\n"
    )

    dic = {}

    messages = []

    for option, data in results.items():
        if option not in dic:
            dic[option] = []

        message += f"{option}: {data['total_weight']} голосов\n"
        if data["voters"]:
            message += "Участники:\n" + "\n".join(data["voters"]) + "\n\n"
        else:
            message += "0\n\n"

        messages.append(message)
    print(f"messages = {messages}")
    print(f"data = {data}")
    print(f"results = {results}")
    print(f"dic = {dic}")
    print(poll_data[poll_id]['question'])
    print(dic.keys())

    ww_options = dic.keys()
    ww_question = poll_data[poll_id]['question']
    ww_answers = []
    jj = 1
    for i in results:
        ite = results[i]

        aa = str(ite["total_weight"])
        if len(ite["voters"]) >0:
            aa += " ("
            nfirst = False
            for l in ite["voters"]:
                if nfirst:
                    aa += ", "
                aa += l
                nfirst = True
            aa+= ")"

        ww_answers.append(aa)






    # Добавление текущей даты и времени
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    message += f"Отчет сгенерирован: {current_time}"

    await query.message.reply_text(message)

    # Создание пустого Word файла
    doc = create_hse_docx(2025, 1, 27, "**ДАТА** 2025 года", total_weight, total_voted_weight, council_members_count, total_voters_count, "___ИМЯ___", ww_options, ww_answers, ww_question)
    #doc.add_heading("Отчет по голосованию", level=1)
    #doc.add_paragraph(message)

    # Сохранение файла в память
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    # Отправка файла пользователю
    await query.message.reply_document(document=buffer, filename="voting_report.docx")

# Основной код
if __name__ == "__main__":
    application = ApplicationBuilder().token("8315966318:AAFZ4VSBJZXBnQ1iNPfpbmOqFcuJ5jxeYb8").build()

    application.add_handler(CommandHandler("start", start))
    application.add_handler(PollAnswerHandler(handle_poll_answer))
    application.add_handler(CallbackQueryHandler(create_test_poll, pattern="create_test_poll"))
    application.add_handler(CallbackQueryHandler(view_results, pattern="view_results"))
    application.add_handler(CallbackQueryHandler(show_voters, pattern="view_poll_"))
    application.add_handler(CallbackQueryHandler(start_custom_poll, pattern="create_poll"))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_custom_poll))
    #application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_custom_poll_options))


    print("Бот запущен...")
    application.run_polling()