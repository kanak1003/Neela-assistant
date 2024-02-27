import speech_recognition as sr
import pyttsx3
import os
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import pyautogui
import webbrowser
import time
from plyer import notification
import datetime


def get_engine():
    if not hasattr(get_engine, "_engine"):
        get_engine._engine = pyttsx3.init()
        voices = get_engine._engine.getProperty('voices')
        get_engine._engine.setProperty('voice', voices[1].id)  # Use a female voice
    return get_engine._engine

def speak(text):
    engine = get_engine()
    engine.say(text)
    engine.runAndWait()

def play_yes_notification():
    speak("Yes")
    notification.notify(
        title="Voice Assistant Activated",
        message="How can I assist you?",
        app_name="Neela",
    )
def get_time_of_day():
    current_hour = datetime.datetime.now().hour
    if 5 <= current_hour < 12:
        return "morning"
    elif 12 <= current_hour < 14:
        return "noon"
    elif 14 <= current_hour < 17:
        return "noon"
    elif 17 <= current_hour < 21:
        return "evening"
    else:
        return "night"
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        try:
            print("Listening...")
            audio = r.listen(source)
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except sr.UnknownValueError:
            return "Sorry, I didn't get that. Can you please repeat?"
        except sr.RequestError as e:
            return f"Error during speech recognition: {e}"
def process_command(command):

    
    if "hello" in command:
        return "Hello! How can I assist you today?"
    elif "what is your name" in command:
        return "I am your voice assistant."
    elif "open file explorer" in command:
        os.system("explorer")
        return "Opening File Explorer."
    elif "close file explorer" in command:
        os.system("taskkill /f /im explorer.exe")
        return "Closing File Explorer."
    elif "open this pc" in command:
        os.system("explorer /e,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}")
        return "Opening This PC in File Explorer."
    elif "close recent window" in command:
        pyautogui.hotkey("alt", "f4")
        return "Closing recent window."
    elif "open chrome" in command:
        os.system("start chrome")
        return "Opening Google Chrome."
    elif "fuck" in command or "fuck you" in command:
        speak("You can't because Your penis is too small")
        return()
    elif "close chrome" in command:
        os.system("taskkill /f /im chrome.exe")
        return "Closing Google Chrome."
    elif "open photoshop" in command:
        os.startfile('"E:\\Adobe Photoshop (Beta)\\Photoshop.exe"')
        return "Opening Photo shop."
    elif "close photoshop" in command:
        os.system("taskkill /f /im Photoshop.exe")
        return "Closing Photoshop."
    elif "open steam" in command:
        os.startfile('"C:\\Program Files (x86)\\Steam\\Steam.exe"')
        return "Opening steam."
    elif "close steam" in command:
        os.system("taskkill /f /im Steam.exe")
        return "Closing steam."

    elif "open youtube" in command:
        webbrowser.open("https://www.youtube.com")
        return "Opening YouTube in Chrome."
    elif "open facebook" in command:
        webbrowser.open("https://www.facebook.com")
        return "Opening Facebook in Chrome."
    elif "search" in command:
        # Extract the search query from the command
        search_query = command.replace("search", "").strip()
        
        # Open Chrome (if not already open)
        os.system("start chrome")
        time.sleep(2)  # Wait for Chrome to open

        # Simulate keyboard input for the search query
        pyautogui.typewrite(search_query)
        pyautogui.press('enter')
        return f"Searching for {search_query} on Google."

    elif "shutdown" in command:
        os.system("shutdown /s /t 1")
        return "Shutting down the computer and good bye see you soon."
    # Change the variable name from 'min' to something else


    elif "play video" in command or "stop video" in command:
        pyautogui.hotkey("k")  # Press 'k' to play/pause video
        return "Playing/stop the video."

    elif "next video" in command:
        pyautogui.hotkey("shift", "n")  # Press Shift + 'n' for next video
        return "Playing the next video."

    elif "previous video" in command:
        pyautogui.hotkey("shift", "p")  # Press Shift + 'p' for previous video
        return "Playing the previous video."

    elif "full screen" in command:
        pyautogui.hotkey("f")  # Press 'f' for full screen
        return "Entering full screen mode."
    elif "close all windows and shutdown" in command:
        # Use Alt + F4 to close the active window (simulates pressing Alt + F4)
        pyautogui.hotkey("alt", "f4") and os.system("shutdown /s /t 1") 
        return "Closing all windows."
    elif "what time is it" in command or "tell me the time" in command:
        hour = datetime.datetime.now().strftime("%H")
        min = datetime.datetime.now().strftime("%M")
        return f"The current time is {hour} and {min}"
    elif "set volume" in command:
        try:
           # Extract the volume level from the command
            volume_level = int(command.split()[-1])
            # Ensure the volume level is within the valid range (0 to 100)
            volume_level = max(0, __builtins__.min(100, volume_level))
            # Set the system volume
            set_system_volume(volume_level)
            return f"Setting volume to {volume_level}."
        except ValueError:
            return "Invalid volume level. Please provide a number between 0 and 100."


    

def set_system_volume(volume):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None
    ).QueryInterface(IAudioEndpointVolume)
    interface.SetMasterVolumeLevelScalar(volume / 100, None)



def listen_for_wake_word():
    recognizer = sr.Recognizer()

    with sr.Microphone() as source:
        print("Listening for wake word...")
        recognizer.adjust_for_ambient_noise(source)
        try:
            audio = recognizer.listen(source, timeout=5)  # Set a timeout value (in seconds)
        except sr.WaitTimeoutError:
            print("Timeout. Please try again.")
            return False

    try:
        print("Recognizing wake word...")
        wake_word = recognizer.recognize_google(audio).lower()
        if "neela" in wake_word:
            play_yes_notification()  # Say "Yes" and show notification
            return True
        else:
            return False
    except sr.UnknownValueError:
        print("Could not recognize wake word. Please try again.")
        return False
    except sr.RequestError as e:
        print(f"Could not request results from Google Speech Recognition service; {e}")
        return False

def main():
    print("Welcome to Neela - Your Voice Assistant!")
    time_of_day = get_time_of_day()
    speak(f"Good {time_of_day}!Welcome back to Neela, How can I assist you today?")

    while True:
        if listen_for_wake_word():
            recognizer = sr.Recognizer()

            with sr.Microphone() as source:
                print("Listening...")
                recognizer.adjust_for_ambient_noise(source)
                audio = recognizer.listen(source)

            try:
                print("Recognizing...")
                command = recognizer.recognize_google(audio).lower()
                print("You said:", command)

                response = process_command(command)
                speak(response)
                print("Response:", response)

            except sr.UnknownValueError:
                print("Sorry, I could not understand your audio.")
            except sr.RequestError as e:
                print(f"Could not request results from Google Speech Recognition service; {e}")

if __name__ == "__main__":
    main()

