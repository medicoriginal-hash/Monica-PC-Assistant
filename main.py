#BY MEDIC
import speech_recognition as sr
import pyttsx3
import webbrowser
import customtkinter as ctk
import json
import threading
import time
import os
import subprocess
import pyautogui
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import audioop
from threading import Lock
import pyperclip
from PIL import Image
from datetime import datetime

#loading images
imagepath_announce = os.path.join("assets", "announce.png")
image_announce = ctk.CTkImage(Image.open(imagepath_announce), size=(380, 300))

imagepath_default = os.path.join("assets", "default.png")
image_default = ctk.CTkImage(Image.open(imagepath_default), size=(250, 250))

imagepath_listen = os.path.join("assets", "listen.png")
image_listen = ctk.CTkImage(Image.open(imagepath_listen), size=(250, 250))

imagepath_happy = os.path.join("assets", "happy.png")
image_happy = ctk.CTkImage(Image.open(imagepath_happy), size=(250, 250))

mic_lock = Lock()

# gui init
engine = pyttsx3.init()
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# json loading
with open("commands.json", "r", encoding="utf-8") as f:
    commands = json.load(f)

    paths = json.load(open("paths.json"))

    
# audio
def speak(text):
    engine.say(text)
    engine.runAndWait()

#time checking
def format_time():
    now = datetime.now()
    hours = now.hour
    minutes = now.minute

    h_form = (
        "час" if hours % 10 == 1 and hours % 100 != 11 else
        "часа" if 2 <= hours % 10 <= 4 and not 12 <= hours % 100 <= 14 else
        "часов"
    )

    m_form = (
        "минута" if minutes % 10 == 1 and minutes % 100 != 11 else
        "минуты" if 2 <= minutes % 10 <= 4 and not 12 <= minutes % 100 <= 14 else
        "минут"
    )

    return f"{hours} {h_form}, {minutes} {m_form}"


# proccesing commands
def process_command(command, listening_forever=False):
    command = command.lower()

    # emotes
    if any(phrase in command for phrase in ["как настроение"]):
        emotion_label.configure(image=image_happy)
        speak("У меня всё отлично, спасибо!")
        return
    for key in commands:
        if key in command:
            action = commands[key]
            if action == "open_browser":
                webbrowser.open("https://google.com")
                log("Открыт браузер")
                speak("Открываю браузер")
            elif action == "open_youtube":
                webbrowser.open("https://youtube.com")
                log("Открыт YouTube")
                speak("Открываю ютуб")
            elif action == "open_vk":
                webbrowser.open("https://vk.com")
                log("Открыт ВКонтакте")
                speak("Открываю ВКонтакте")
            elif action == "open_steam":
                subprocess.Popen(paths["steam"])
                log("Открыт Steam")
                speak("Открываю стим")
            elif action == "open_discord":
                subprocess.Popen([paths["discord"], "--processStart", "Discord.exe"])
                log("Открыт Discord")
                speak("Открываю дискорд")
            elif action == "shutdown":
                speak("Выключаю компьютер")
                os.system("shutdown /s /t 5")
            elif action == "restart":
                speak("Перезагружаю компьютер")
                os.system("shutdown /r /t 5")
            elif action == "open_cmd":
                subprocess.Popen("cmd.exe")
                log("Открыт командный интерпретатор")
                speak("Открываю командную строку")
            elif action == "open_folder":
                os.startfile("C:\\")
                log("Открыт диск C")
                speak("Открываю диск C")
            elif action == "volume_up":
                set_volume(10)
                speak("Громкость увеличена")
            elif action == "volume_down":
                set_volume(-10)
                speak("Громкость уменьшена")
            elif action == "pause_play":
                pyautogui.press("playpause")
                speak("Пауза")
            elif action == "minimize_window":
                pyautogui.hotkey("win", "down")
                speak("Окно свернуто")
            elif action == "fullscreen":
                pyautogui.press("f")
                speak("Полноэкранный режим")
            elif action == "windowed":
                pyautogui.press("f11")
                speak("Оконный режим")
            elif action == "open_roblox":
                webbrowser.open("https://www.roblox.com/home")
                log("Открыт Роблокс")
                speak("Роблокс открыт")
            elif action == "open_gpt":
                webbrowser.open("https://chatgpt.com")
                log("Открыт ЧатГПТ")
                speak("ЧатГПТ открыт")
            elif action == "close_tab":
                pyautogui.hotkey("ctrl", "w")
                log("Вкладка закрыта")
                speak("Закрыла вкладку")
            elif action == "open_mail":
                webbrowser.open("https://mail.ru")
                log("Открыта почта")
                speak("Почта открыта")
            elif action == "search":
                search_query = command.replace("поиск", "", 1).replace("найди", "", 1).replace("поищи", "", 1).strip()
                if search_query:
                    webbrowser.open(f"https://www.google.com/search?q={search_query}")
                    log(f"Поиск: {search_query}")
                    speak(f"Ищу {search_query}")
                else:
                    speak("Пожалуйста, уточните, что искать.")              
            elif action == "close_window":
                pyautogui.hotkey('alt', 'f4')
                speak("Закрыла окно")
            elif action == "click":
                pyautogui.click()
            elif action == "next":
                pyautogui.press("nexttrack")
                log("Переключено")
                speak("Переключила")
            elif action == "open_telegram":
                subprocess.Popen([paths["telegram"]])
                log("Открыт Телеграм")
                speak("Открываю Телеграм")
# управление мышкой
            elif action == "mouse_up":
                pyautogui.move(0, -100, duration=0.5)
                log("Курсор передвинут вверх")
            elif action == "mouse_down":
                pyautogui.move(0, 100, duration=0.5)
                log("Курсор передвинут вниз")
            elif action == "mouse_right":
                pyautogui.move(100, 0, duration=0.5)
                log("Курсор передвинут вправо")
            elif action == "mouse_left":
                pyautogui.move(-100, 0, duration=0.5)
                log("Курсор передвинут влево")
# конец управления мышкой (шоб не потеряться)
            elif action == "type":
                text_to_type = command.replace("напиши", "").replace("пиши", "").replace("напиши", "").strip()
                if text_to_type:
                 pyperclip.copy(text_to_type)
                pyautogui.hotkey('ctrl', 'v')
                pyperclip.paste()
                speak("Готово")
            elif action == "enter":
                pyautogui.press("enter")
                log("Enter нажат")      
                speak("Потвердила")      
            elif action == "next_tab":
                pyautogui.hotkey("ctrl","tab")
                log("Вкладка переключена")
                speak("Переключила")    
            elif action == "open_music":
                webbrowser.open("https://vk.com/audios811352879")
                log("Музыка открыта")
                speak("Открыла музыку")          
            elif action == "open_github":
                webbrowser.open("https://github.com/trending")
                log("Гитхаб открыт")
                speak("Открываю Гитхаб")         
            elif action == "timecheck":
                log(format_time())
                speak(format_time())            
            else:
                log(f"Неизвестное действие: {action}")
            return
    if not listening_forever:
        log("Команда не распознана")
        speak("Повторите пожалуйста")

# audio holder
def set_volume(change):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    current = volume.GetMasterVolumeLevelScalar()
    new_volume = min(1.0, max(0.0, current + (change / 100)))
    volume.SetMasterVolumeLevelScalar(new_volume, None)

# recognize speech
recognizer = sr.Recognizer()
selected_mic_index = None  

# getting microphones 
def get_microphone_list():
    raw_names = sr.Microphone.list_microphone_names()
    fixed_names = []
    for name in raw_names:
        try:
            fixed = name.encode('latin1').decode('utf-8')
            fixed_names.append(fixed)
        except:
            fixed_names.append(name)
    return fixed_names

# voice indicatod (WON'T WORK)
def update_voice_meter():
    while True:
        try:
            with mic_lock:
                with sr.Microphone(device_index=selected_mic_index) as source:
                    recognizer.adjust_for_ambient_noise(source, duration=0.2)
                    while True:
                        try:
                            audio = recognizer.listen(source, timeout=0.3, phrase_time_limit=0.1)
                            audio_data = audio.get_raw_data()

                            rms = audioop.rms(audio_data, 2)  
                            level = min(100, rms / 32767 * 200)  

                            app.after(0, lambda: voice_meter.set(level/100))
                        except sr.WaitTimeoutError:
                            app.after(0, lambda: voice_meter.set(0))
                        time.sleep(0.05)
        except OSError as e:
            app.after(0, lambda: log(f"Ошибка доступа к микрофону: {str(e)}"))
            time.sleep(2)
        except Exception as e:
            app.after(0, lambda: log(f"Ошибка индикатора: {str(e)}"))
            break
listen_forever = False

# listen func
def listen():
    emotion_label.configure(image=image_listen)
    log("Слушаю...")
    try:
        with sr.Microphone(device_index=selected_mic_index) as source:
            recognizer.adjust_for_ambient_noise(source)
            while listen_forever:
                audio = recognizer.listen(source, timeout=None)  
                query = recognizer.recognize_google(audio, language="ru-RU")
                log(f"Вы сказали: {query}")
                process_command(query, listening_forever=True)
                time.sleep(0.5)
    except sr.UnknownValueError:
        pass 
    except sr.RequestError:
        log("Ошибка запроса к сервису распознавания")
        speak("Проблема с интернетом")
    except Exception as e:
        log(f"Ошибка: {e}")
        speak("Произошла ошибка")
    finally:
        emotion_label.configure(image=image_default)

# listening forever

def toggle_listen_forever():
    global listen_forever
    listen_forever = not listen_forever
    if listen_forever:
        listen_forever_button.configure(text="Слушать всегда: Включено")
        threading.Thread(target=start_listening_forever, daemon=True).start()
    else:
        listen_forever_button.configure(text="Слушать всегда: Выключено")

def start_listening_forever():
    while listen_forever:
        listen()

# main gui
app = ctk.CTk()
app.title("Monica : PC Assistant")
app.geometry("500x780")

#emotion_label = ctk.CTkLabel(app, text=":|", font=("Arial", 40)) #old emotions
#emotion_label.pack(pady=5)

emotion_label = ctk.CTkLabel(app, image=image_default, text="")
emotion_label.pack(pady=1)

# popup show
popup = ctk.CTkToplevel()
popup.title("Внимание")
popup.geometry("480x400")
ctk.CTkLabel(popup, image=image_announce, text="").pack(pady=10)
ctk.CTkLabel(popup, text="Помните! это очень ранний доступ и все может поменяться!", font=("Arial", 14)).pack(pady=5)
ctk.CTkLabel(popup, text="V0.13", font=("Arial", 18)).pack(pady=6)
ctk.CTkButton(popup, text="ОК", command=popup.destroy).pack(pady=10)

logbox = ctk.CTkTextbox(app, width=480, height=300)
logbox.pack(pady=10)
#🐟
mic_names = get_microphone_list()
microphone_menu = ctk.CTkOptionMenu(app, values=["По умолчанию"] + mic_names, command=lambda name: set_microphone(name))
microphone_menu.set("По умолчанию")
microphone_menu.pack(pady=10)

listen_button = ctk.CTkButton(app, text="🎙 Слушать команду", command=lambda: threading.Thread(target=listen).start())
listen_button.pack(pady=10)

listen_forever_button = ctk.CTkButton(app, text="Слушать всегда: Выключено", command=toggle_listen_forever)
listen_forever_button.pack(pady=10)

voice_meter = ctk.CTkProgressBar(app, width=400)
voice_meter.set(0)
voice_meter.pack(pady=10)

def set_microphone(name):
    global selected_mic_index
    if name == "По умолчанию":
        selected_mic_index = None
        log("Выбран микрофон по умолчанию ОС")
    else:
        try:
            selected_mic_index = mic_names.index(name)
            log(f"Выбран микрофон: {name}")
        except ValueError:
            log("Ошибка выбора микрофона")

def log(text):
    logbox.insert("end", text + "\n")
    logbox.see("end")
    print(text)

speak("К вашим услугам!")
log("Моника готова к работе")
log("Выберите микрофон из списка выше или оставьте по умолчанию")

threading.Thread(target=update_voice_meter, daemon=True).start()

app.mainloop()
