import datetime
from itertools import permutations
from datetime import time
from flask import Flask, render_template, request, jsonify, redirect, session
import io
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google_auth_oauthlib.flow import Flow
import os
import googleapiclient.discovery
import google.oauth2.credentials
import pandas as pd
import json

app = Flask(__name__)
# Комплексні пропозиції
COMPLEX_OFFERS = {
    "MINI": [("Arena", 60, "ArenaText"), ("LL", 20, "LLText")],
    "STANDART": [("Arena", 60, "ArenaText"), ("Kvest", 50, "KvestText"), ("LL", 20, "LLText")],
    "MAXI": [("Arena", 60, "ArenaText"), ("Kvest", 50, "KvestText"), ("LL", 60, "LLText")],
    "MEGA": [("Arena", 60, "ArenaText"), ("Kvest", 60, "KvestText"), ("ArenaActor", 30, "ArenaAText"), ("LL", 60, "LLText")],
}   

app.secret_key = "my_super_secret_key_123"  # постійний, а не random
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'

# Налаштування OAuth
# Зчитування секрету з Render
CLIENT_SECRETS_JSON = os.environ.get("GOOGLE_CREDENTIALS")
CLIENT_SECRETS_DICT = json.loads(CLIENT_SECRETS_JSON)

SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
REDIRECT_URI = 'http://localhost:5000/oauth2callback'


@app.route('/') 
#---Головна сторінка з посиланням на авторизацію через Google---(1p)
def index():    
    return render_template('Logo.html')


@app.route('/authorize')
def authorize():
    flow = Flow.from_client_config(
        CLIENT_SECRETS_DICT,  # <-- словник, а не шлях до файлу
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI
    )
    auth_url, state = flow.authorization_url(
        access_type='offline',
        include_granted_scopes='true'
    )
    session['state'] = state
    return redirect(auth_url)

@app.route('/oauth2callback')
def oauth2callback():
    if 'state' not in session:
        return redirect('/authorize')
    state = session['state']
    flow = Flow.from_client_config(
        CLIENT_SECRETS_DICT,  # <-- теж заміна
        scopes=SCOPES,
        state=state,
        redirect_uri=REDIRECT_URI
    )
    flow.fetch_token(authorization_response=request.url)
    credentials = flow.credentials
    session['credentials'] = credentials_to_dict(credentials)
    return redirect('/list_files')

@app.route("/list_files", methods=["GET", "POST"])
def list_files():
    if "credentials" not in session:
        return redirect("authorize")

    credentials = google.oauth2.credentials.Credentials(**session["credentials"])
    drive_service = googleapiclient.discovery.build("drive", "v3", credentials=credentials)

    results = drive_service.files().list(
        pageSize=100, fields="files(id, name, webViewLink)"
    ).execute()
    items = results.get("files", [])

    return render_template("list_files.html", files=items)




@app.route('/load_excel', methods=['POST'])
#---"""Завантажує Excel-файл з Google Drive у пам'ять та повертає кількість рядків"""---(3p)
def load_excel():
    file_id = request.json['fileId']
    creds = get_credentials_from_session()
    if not creds:
        return jsonify({"error": "Не авторизовано"}), 401

    service = build('drive', 'v3', credentials=creds)

    request_drive = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request_drive)
    done = False
    while done is False:
        status, done = downloader.next_chunk()

    fh.seek(0)
    df = pd.read_excel(fh)  # pandas читає файл напряму з пам’яті
    # далі робиш свою обробку даних
    return jsonify({"rows": len(df)})



# ======= ДОПОМІЖНІ ФУНКЦІЇ =======

def credentials_to_dict(credentials):
    return {
        'token': credentials.token,
        'refresh_token': credentials.refresh_token,
        'token_uri': credentials.token_uri,
        'client_id': credentials.client_id,
        'client_secret': credentials.client_secret,
        'scopes': credentials.scopes
    }

from google.oauth2.credentials import Credentials

def google_auth_from_dict(creds_dict):
    return Credentials(
        token=creds_dict['token'],
        refresh_token=creds_dict.get('refresh_token'),
        token_uri=creds_dict['token_uri'],
        client_id=creds_dict['client_id'],
        client_secret=creds_dict['client_secret'],
        scopes=creds_dict['scopes']
    )

def get_credentials_from_session():
    if "credentials" not in session:
        return None
    creds_dict = session["credentials"]
    return Credentials(
        token=creds_dict['token'],
        refresh_token=creds_dict.get('refresh_token'),
        token_uri=creds_dict['token_uri'],
        client_id=creds_dict['client_id'],
        client_secret=creds_dict['client_secret'],
        scopes=creds_dict['scopes']
    )



def time_to_minutes(t): #Перетворює час (datetime.time) у хвилини.
    return t.hour * 60 + t.minute

def minutes_to_time(minutes):   #Перетворює кількість хвилин у формат datetime.time.
    hours = minutes // 60
    minutes = minutes % 60
    return time(hours, minutes)

def get_activity_text(activity, C_value):  #Назву гри залежно від вхідного коду C_val
    C_value = C_value.upper()
    if activity == "Arena":
        if "ЛТ" in C_value:
            return "Лт"
        elif "NERF" in C_value:
            return "Nerf"
        elif "АРЕНА" in C_value:
            return "Арена"
        elif "НЕРФ" in C_value:
            return "Нерф"
        elif ("Фредді" or "Фреді") in C_value:
            return "Фредді"
        else:
            return "Арена"
    elif activity == "LL":
        if "ЛЛ" in C_value:
            return "ЛЛ"
        elif "ТИР" in C_value:
            return "Тир"
        elif "ЛЛ/ТИР" in C_value or 'ТИР/ЛЛ' in C_value:
            return "ЛЛ/ТИР"
        else:
            return "ЛЛ"
    elif activity == "Kvest":
        if "КВЕСТ" in C_value:
            return "Квест"
        elif "МІСІЯ" in C_value:
            return "Місія"
        elif "СТАЛКЕР" in C_value:
            return "Сталкер"
        elif "ГАРРІ ПОТТЕР" in C_value:
            return "Гаррі Поттер"
        elif "ГП" in C_value:
            return "ГП"
        else:
            return "Квест"
    elif activity == "ArenaActor":
        if "ЛТ(А)" in C_value:
            return "ArenaAText"
        elif "А" in C_value:
            return "Гра з Актором"
        else:
            return "Гра з Актором"
    return ""

schedule = []
def get_schedule(offer, start_time, C_value, order=None):  # Формує розклад для комплексної пропозиції.
    if offer == "MINI":
        if "КВЕСТ" in C_value or "СТАЛКЕР" in C_value or "ГП" in C_value or "МІСІЯ" in C_value:
            a = "Kvest"
        elif "Арена" in C_value or "ЛТ" in C_value or "НЕРФ" in C_value or "ЛТ/НЕРФ" in C_value:
            a = "Arena"
        else:
            a = ""
        
        # Оновлюємо COMPLEX_OFFERS для MINI з врахуванням значення a
        COMPLEX_OFFERS[offer] = [(a, 60, "ArenaText"), ("LL", 20, "LLText")]
    
    global schedule 
    schedule = []
    current_time = time_to_minutes(start_time)
    games = COMPLEX_OFFERS[offer]
    if order:
        games = [games[i] for i in order]

    for i, (game, duration, name) in enumerate(games):
        name = get_activity_text(game, C_value)
        start = current_time
        end_time = start + duration
        schedule.append((game, minutes_to_time(start), minutes_to_time(end_time), name))
        print(schedule)

        current_time = end_time
        if i == 0:
            current_time += 20
        else:
            current_time += 10
    
    
    return schedule



def check_conflict(schedule1, schedules):  # Перевіряє, чи є конфлікт між розкладами.
    for schedule2 in schedules:
        for game1, start1, end1, _ in schedule1:
            for game2, start2, end2, _ in schedule2:
                if (game1 == game2 or (game1, game2) in [("Arena", "ArenaActor"), ("ArenaActor", "Arena")]) and start1 < end2 and end1 > start2:
                    return True
    return False

def split_and_schedule_games(C_value, start_time): # Розділяє рядок C_value з некомплексними іграми і формує розклад.
    games = C_value.split('+')
    global schedule
    schedule = []
    current_time = time_to_minutes(start_time)

    
    for game in games:
        if len(game) < 3 or not game[2:].isdigit():  # Перевірка формату гри
            print(f"Помилка: Невірний формат гри '{game}' у C_value '{C_value}'. Пропуск.")
            continue
        
        name, duration = game[:2], int(game[2:])
        end_time = current_time + duration
        name_lower = name.lower()  # Приводимо до нижнього регістру для порівняння

        if name_lower in {"arena", "лт", "nerf", "нєрф", "арена"}:
            game = "Arena"
        elif name_lower in {"ll", "лл", "тир", "лл/тир", "тир/лл"}:
            game = "LL"
        elif name_lower in {"kvest", "квест", "гп", "гаррі поттер", "сталкер", "місія"}:
            game = "Kvest"
        elif name_lower in {"arenaactor", "лт(а)", "а", "гра з актором"}:
            game = "ArenaActor"
        else:
            print(f"Помилка: Невідома гра '{name}' у C_value '{C_value}'. Пропуск.")
            continue

        schedule.append((game, minutes_to_time(current_time), minutes_to_time(end_time), name))
        current_time = end_time + 10  

    return schedule

def rearrange_schedule(offer, new_schedule, all_schedules, start_time, C_value): # Пробує знайти альтернативне розташування ігор, щоб уникнути конфліктів.
    if offer in COMPLEX_OFFERS:
        games_count = len(COMPLEX_OFFERS[offer])
    else:
        games = C_value.split('+')  # Розділення ігор для некомплексної пропозиції
        games_count = len(games)

    for order in permutations(range(games_count)):
        if offer in COMPLEX_OFFERS:
            new_schedule = get_schedule(offer, start_time, C_value, order)
        else:
            # Формуємо новий порядок для некомплексних ігор
            new_games = '+'.join(games[i] for i in order)
            new_schedule = split_and_schedule_games(new_games, start_time)

        if not check_conflict(new_schedule, all_schedules):
            return new_schedule
    return None 
 
all_schedules = []


def read_excel_from_google_drive(file_id):
    credentials = get_credentials_from_session()
    if not credentials:
        print("Не авторизовано")
        return None

    drive_service = build("drive", "v3", credentials=credentials)
    try:
        # Для Google Sheets замість get_media -> export
        request = drive_service.files().export_media(
            fileId=file_id,
            mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'  # Excel формат
        )
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            if status:
                print(f"Завантажено {int(status.progress() * 100)}%")
        fh.seek(0)
        df = pd.read_excel(fh, engine='openpyxl')
        print("Файл успішно зчитано")
        return df
    except Exception as e:
        print("Помилка при завантаженні або зчитуванні файлу:", e)
        return None


from datetime import datetime, time

def parse_excel_time(value):
    if isinstance(value, time):
        return value  # вже datetime.time
    elif isinstance(value, datetime):
        return value.time()  # datetime -> time
    elif isinstance(value, str):
        try:
            hour, minute = map(int, value.split(':')[:2])
            return time(hour, minute)
        except:
            return None
    elif isinstance(value, (int, float)):
        # Excel часто зберігає час як дробове число дні
        hours = int(value * 24)
        minutes = int((value * 24 - hours) * 60)
        return time(hours, minutes)
    return None

   

@app.route('/schedule', methods=['POST'])
def schedule():
    file_id = request.form.get("file_id")
    if not file_id:
        return "Помилка: не вказаний file_id", 400

    df = read_excel_from_google_drive(file_id)
    if df is None:
        return "Помилка: не вдалося зчитати Excel з Google Drive", 400

    data = df.values.tolist()

    if not data:
        print("Помилка: файл Excel порожній або дані не знайдені.")
        return "Помилка: файл Excel порожній або дані не знайдені.", 400

    # --- додаємо вибір дня (наприклад, передається з форми) ---
    selected_day = request.form.get("day")  # очікуємо рядок '2025-09-2.7'
    if not selected_day:
        return "Помилка: не вибрано день", 400

    try:
        selected_day = pd.to_datetime(selected_day).date()  # приводимо до дати
    except:
        return "Неправильний формат дати", 400

    result = []
    global all_schedules
    all_schedules = []

    for row in data:
        try:
            # --- тут беремо дату з Excel ---
            excel_date = pd.to_datetime(row[1]).date()

            # якщо день не співпадає — пропускаємо рядок
            if excel_date != selected_day:
                continue
        except Exception as e:
            print("Помилка з датою:", e, row)
            continue

        complex_offer = str(row[4]).strip().upper()
        start_time_str = row[3]
        extra_info = row[5]

        try:
            start_time = parse_excel_time(start_time_str)
            if start_time is None:
                    print(f"Недійсний формат часу в Excel для рядка '{row}'.")
                    continue

        except ValueError:
            print(f"Недійсний формат часу в Excel для рядка '{row}'.")
            continue

        if complex_offer in COMPLEX_OFFERS:
            if complex_offer == "MINI" or complex_offer == "STANDART" or complex_offer == "MAXI" or complex_offer == "MEGA" :
                C_value = str(extra_info).strip().upper()
            else:
                C_value = "" 

            new_schedule = get_schedule(complex_offer, start_time, C_value)
            if check_conflict(new_schedule, all_schedules):
                rearranged_schedule = rearrange_schedule(complex_offer, new_schedule, all_schedules, start_time, C_value)
                if rearranged_schedule:
                    all_schedules.append(rearranged_schedule)
                    result.append({"complex_offer": complex_offer, "schedule": rearranged_schedule})
                else:
                    print(f"Неможливо переставити ігри. Виберіть інший час для КП - {complex_offer}")
            else:
                all_schedules.append(new_schedule)
                result.append({"complex_offer": complex_offer, "schedule": new_schedule})

            # Вивід часу після виведення розкладу
            if complex_offer == "MINI":
                end_time = minutes_to_time(time_to_minutes(start_time) + 120)
            elif complex_offer == "STANDART":
                end_time = minutes_to_time(time_to_minutes(start_time) + 180)
            elif complex_offer == "MAXI":
                end_time = minutes_to_time(time_to_minutes(start_time) + 240)
            elif complex_offer == "MEGA":
                end_time = minutes_to_time(time_to_minutes(start_time) + 300)
            
            result[-1]["end_time"] = end_time
            
        else:
    # Логіка для некомплексних пропозицій
            C_value = str(extra_info).strip().upper()
            new_schedule = split_and_schedule_games(C_value, start_time)
            if check_conflict(new_schedule, all_schedules):
                print("Конфлікт")
                rearranged_schedule = rearrange_schedule(complex_offer, new_schedule, all_schedules, start_time, C_value)
                print("Конфлікт")
                if rearranged_schedule:
                    all_schedules.append(rearranged_schedule)
                    result.append({"complex_offer": "Некомплексна пропозиція", "schedule": rearranged_schedule})
                else:
                    print("Неможливо переставити ігри. Виберіть інший час для ігор:", C_value)
            else:
                all_schedules.append(new_schedule)
                result.append({"complex_offer": "Некомплексна пропозиція", "schedule": new_schedule})
        
 
    return render_template('schedule.html', schedules=result )

@app.route('/fetch_schedule/<offer_type>')
def fetch_schedule(offer_type):
    
    global all_schedules
    
    if offer_type in ["STANDART", "MAXI", "MEGA"]:
        OFFER_DURATION = sum([game[1] for game in COMPLEX_OFFERS[offer_type]])  # Тривалість у хвилинах
        OFFER_INTERVAL = (10 * 60, 21 * 60)  # Години в хвилинах
        interval = 30  # Інтервал у хвилинах між спробами
        inserted_schedules = []  # Список для вставлених розкладів

        for start_time in range(OFFER_INTERVAL[0], OFFER_INTERVAL[1] - OFFER_DURATION + 1, interval):
            start_time_obj = minutes_to_time(start_time)

            # Створюємо можливі перестановки ігор з обраної пропозиції
            games = COMPLEX_OFFERS[offer_type]
            for order in permutations(range(len(games))):
                new_schedule = get_schedule(offer_type, start_time_obj, C_value="", order=order)

                # Перевірка на конфлікти
                if not check_conflict(new_schedule, all_schedules):
                    inserted_schedules.append(new_schedule)
                    print(f"{offer_type} успішно вставлено з {start_time_obj.strftime('%H:%M')} в порядку {order}")
                    break  # Якщо успішно вставлено, переходимо до наступного часу

        formatted_schedules = []
        for schedule in inserted_schedules:
            for game, start, end, name in schedule:
                formatted_schedules.append({
                    "game": game,
                    "start_time": start.strftime("%H:%M"),
                    "end_time": end.strftime("%H:%M"),
                    "name": name
                    
                })
        #print(inserted_schedules)
        return jsonify(formatted_schedules)
    return jsonify([]), 404


if __name__ == "__main__":
    app.run(debug=True)
