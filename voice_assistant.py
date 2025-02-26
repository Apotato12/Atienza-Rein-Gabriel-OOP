import os
import speech_recognition
import pyttsx3
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
import pyautogui

# Starts the text to speech engine
def start_py3():
    engine = pyttsx3.init()
    return engine

# Makes the assistant say something
def speak(text):
    engine = start_py3()
    engine.say(text)
    engine.runAndWait()

# Listens to the user's voice
def listen():
    recognizer = speech_recognition.Recognizer()
    with speech_recognition.Microphone() as mic:
        print("Waiting")
        speak("")
        audio = recognizer.listen(mic)
        try:
            command = recognizer.recognize_google(audio)
            print("You said: " + command)
            return command.lower()
        except speech_recognition.UnknownValueError:
            print("Sorry, I didn't hear that.")
            return None
        
def handle_command(command):
    if 'open notepad' in command:
        open_notepad()
    elif 'open chrome' in command:
        open_chrome()
    elif 'open word' in command:
        open_word()
    elif 'open limbus' in command:
        open_limbus()
    elif 'open opera' in command:
        open_opera()
    elif 'close notepad' in command:
        close_app('notepad.exe')
    elif 'close chrome' in command:
        close_app('chrome.exe')
    elif 'close word' in command:
        close_app('WINWORD.EXE')
    elif 'close limbus' in command:
        close_app('LimbusCompany.exe')
    elif 'close opera' in command:
        close_app('Opera.exe')
    elif 'close discord' in command:
        close_app('discord.exe')    
    # Increase volume by 50%
    elif 'volume up' in command:
        change_volume(0.5)
    # Decrease volume by 50%
    elif 'volume down' in command:
        change_volume(-0.5) 
    elif 'switch' in command or 'tab' in command: 
        Switch(command)

#Open applications
def open_notepad():
    os.system('notepad')
    speak("Notepad is opening")

def open_chrome():
    os.startfile('C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe')
    speak("Google Chrome is opening")

def open_word():
    os.startfile('C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE')
    speak("Microsoft Word is opening")

def open_limbus():
    os.startfile('D:\\Steam\\steamapps\\common\\Limbus Company\\LimbusCompany.exe')
    speak("Limbus Company is opening")

def open_opera():
    os.startfile('C:\\Users\\PC\\AppData\\Local\\Programs\\Opera GX\\opera.exe')
    speak("Opening Opera")

# Close applications using os
def close_app(app_name):
    print(f"Attempting to close: {app_name}")
    try:
        os.system(f'taskkill /f /im {app_name}')
        speak(f"{app_name} is closing")
    except:
        speak("I can't close that application.")

# Volume control
def change_volume(change):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = interface.QueryInterface(IAudioEndpointVolume)
    current_volume = volume.GetMasterVolumeLevelScalar()
    new_volume = max(0.0, min(1.0, current_volume + change))  # Volume is between 0.0 and 1.0
    volume.SetMasterVolumeLevelScalar(new_volume, None)
    speak(f"Volume set to {int(new_volume * 100)} percent.")

def Switch(command):
    if "switch" in command or "tab" in command:
        pyautogui.hotkey('alt', 'tab')

def Assistant():
    speak("Hi Sir, I am your voice-controlled assistant. How may I help you?")
    while True:
        command = listen()
        if command:
            handle_command(command)
        if command and 'exit' in command:
            speak("Closing")
            break  

Assistant()