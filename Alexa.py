import speech_recognition as sr
import pyttsx3
import os
import datetime
import subprocess
import time
import pywhatkit
import wikipedia
import screen_brightness_control as sbc
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import pyautogui
import pygame  # Added for sound playback

# Text-to-speech engine
engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
recognizer = sr.Recognizer()

# Contact list
contacts = {
    "sanjay": "+919876543210",
    "vignesh sir": "+919876543210"
}

# Flags
assistant_awake = False
last_response = ""
last_command = ""

def speak(text):
    global last_response
    if text != last_response:
        last_response = text
        print("ALEXA:", text)
        engine.say(text)
        engine.runAndWait()

# New function to play ding sound
def play_ding():
    pygame.mixer.init()
    pygame.mixer.music.load("ding.mp3")
    pygame.mixer.music.play()

def open_software(software_name):
    if 'chrome' in software_name:
        speak('Opening Chrome...')
        subprocess.Popen([r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"])
    elif 'microsoft edge' in software_name:
        speak('Opening Microsoft Edge...')
        subprocess.Popen([r"C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"])
    elif 'notepad' in software_name:
        speak('Opening Notepad...')
        subprocess.Popen(['notepad.exe'])
    elif 'calculator' in software_name:
        speak('Opening Calculator...')
        subprocess.Popen(['calc.exe'])
    elif 'play' in software_name:
        query = software_name.replace('play', '').strip()
        speak(f'Playing {query} on YouTube...')
        subprocess.Popen(["cmd", "/c", f"start https://www.youtube.com/results?search_query={query}"])
    elif 'python portal' in software_name:
        speak('Opening Python Lab Portal...')
        subprocess.Popen([r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe", 
            "https://placement.skcet.ac.in/mycourses/details?id=29938623-5331-4de7-afea-aa9334934454&type=mylabs"])
    elif 'portal' in software_name:
        speak('Opening Sri Krishna Placement Portal...')
        subprocess.Popen([r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe", "https://placement.skcet.ac.in"])
    elif 'i need assistant' in software_name:
        speak('Opening ChatGPT...')
        subprocess.Popen([r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe", "https://chat.openai.com"])
    else:
        speak(f"I couldn't find the software {software_name}")

def close_software(software_name):
    processes = {
        'chrome': "chrome.exe",
        'microsoft edge': "msedge.exe",
        'notepad': "notepad.exe",
        'calculator': "calculator.exe"
    }
    for key, value in processes.items():
        if key in software_name:
            speak(f'Closing {key.capitalize()}...')
            os.system(f"taskkill /f /im {value}")
            return
    speak(f"I couldn't find any open software named {software_name}")

def set_volume(level):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    min_vol, max_vol = volume.GetVolumeRange()[0:2]
    target_volume = min_vol + (max_vol - min_vol) * level
    volume.SetMasterVolumeLevel(target_volume, None)

def set_brightness(level):
    sbc.set_brightness(level)

def send_whatsapp_message(name, message):
    number = contacts.get(name.lower())
    if number:
        speak(f"Sending message to {name}")
        try:
            pywhatkit.sendwhatmsg_instantly(number, message, wait_time=10, tab_close=False)
            time.sleep(2)
            speak(f"Message sent to {name}")
        except Exception as e:
            print("Error:", e)
            speak("Sorry, I couldn't send the message.")
    else:
        speak(f"I don't have a contact named {name}")

def listen():
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source, duration=0.1)
        print("Listening...")
        audio = recognizer.listen(source, phrase_time_limit=5)
    try:
        command = recognizer.recognize_google(audio)
        return command.lower()
    except sr.UnknownValueError:
        return ""
    except sr.RequestError:
        speak("Sorry, the speech service is not available.")
        return ""

def main_loop():
    global assistant_awake, last_command
    while True:
        command = listen().strip()

        if not command or command == last_command:
            continue
        last_command = command

        if not assistant_awake:
            if 'wake up alexa' in command or 'hi alexa' in command:
                assistant_awake = True
                play_ding()
                time.sleep(1)
                speak('Hi sir, I am here, how can I help you?')
            continue

        if 'stop' in command:
            speak('Stopping the program. Say wake up Alexa to start again.')
            assistant_awake = False
            continue

        if 'repeat' in command:
            speak(last_response)

        elif 'open' in command or 'play' in command or 'i need assistant' in command:
            open_software(command)

        elif 'close' in command:
            close_software(command)

        elif 'time' in command:
            current_time = datetime.datetime.now().strftime('%I:%M %p')
            speak(f'The time is {current_time}')

        elif "am i audible" in command:
            speak("Yes sir, I can hear you clearly. How can I help you?")

        elif 'increase volume' in command:
            set_volume(1.0)
            speak('Volume increased')

        elif 'decrease volume' in command:
            set_volume(0.3)
            speak('Volume decreased')

        elif 'set volume to medium' in command:
            set_volume(0.5)
            speak('Volume set to medium')

        elif 'mute volume' in command:
            set_volume(0.0)
            speak('Volume muted')

        elif 'increase brightness' in command:
            set_brightness(100)
            speak('Brightness increased')

        elif 'decrease brightness' in command:
            set_brightness(30)
            speak('Brightness decreased')

        elif 'set brightness to medium' in command:
            set_brightness(50)
            speak('Brightness set to medium')

        elif "tell about our python sir" in command:
            speak("Mister Vignesh is a multi-talented person. He always shares his knowledge with students and is well-known in the Mechatronics Department of Sri Krishna College.")


        elif 'shutdown the system' in command:
            speak('Shutting down the system.')
            os.system('shutdown /s /t 1')

        elif 'restart the system' in command:
            speak('Restarting the system.')
            os.system('shutdown /r /t 1')

        elif 'lock the system' in command:
            speak('Locking the system.')
            os.system('rundll32.exe user32.dll,LockWorkStation')

        elif 'what is your name' in command:
            speak('My name is Alexa, your artificial intelligence.')

        elif 'how can you help me' in command:
            speak('I can assist with your studies, projects, assignments, system control, and more.')

        elif 'send message to' in command:
            found = False
            for name in contacts:
                if name in command:
                    message = command.split(name)[-1].strip()
                    send_whatsapp_message(name, message)
                    found = True
                    break
            if not found:
                speak("I couldn't identify the contact.")

        elif 'who is' in command or 'what is' in command:
            try:
                info = wikipedia.summary(command, sentences=1)
                speak(info)
            except:
                speak("Sorry, I couldn't find information on that.")

if __name__ == "__main__":
    main_loop()
