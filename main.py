from __future__ import with_statement
import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import glob
import datetime
import shutil
import random
import cv2
import pywhatkit as kit
import sys
import pyautogui
import time
import pygetwindow as gw
import operator
import requests
import psutil
import keyboard
import ctypes
import subprocess
import pygetwindow as gw
from pywinauto.application import Application
import win32gui
import win32con

#for friend mode
import chat_friend
is_friend_mode = False
chat_history = []

# Virtual-Key codes
VK_MEDIA_PLAY_PAUSE = 0xB3
VK_MEDIA_NEXT_TRACK = 0xB0
VK_MEDIA_PREV_TRACK = 0xB1
VK_MEDIA_STOP = 0xB2
VK_VOLUME_MUTE = 0xAD
VK_VOLUME_DOWN = 0xAE
VK_VOLUME_UP = 0xAF

last_captured_image = None

def press_key(hexKeyCode):
    ctypes.windll.user32.keybd_event(hexKeyCode, 0, 0, 0)
    ctypes.windll.user32.keybd_event(hexKeyCode, 0, 2, 0)

# Replace with the actual path of chrome.exe on your PC
chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
# OR if you're on 64-bit Windows and Chrome is in Program Files (x86)
# chrome_path = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

# Register 'chrome' browser
webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)
engine.setProperty('rate', 150)

# for index, voice in enumerate(voices):
#     print(f"{index}: {voice.name} ({voice.gender}) - {voice.id}")

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning!")
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")
    speak("What can I do for you ?")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")
        return query  # ✅ return the spoken text
    except Exception as e:
        print("Say that again please...")
        return "None"  # ✅ return real None

last_minimized_hwnd = None  # global state

def minimize_active_window():
    global last_minimized_hwnd
    hwnd = win32gui.GetForegroundWindow()
    if hwnd:
        last_minimized_hwnd = hwnd  # store it
        win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)

def maximize_last_window():
    global last_minimized_hwnd
    if last_minimized_hwnd:
        win32gui.ShowWindow(last_minimized_hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(last_minimized_hwnd)
    else:
        speak("There is no minimized window to maximize.")


if __name__ == "__main__":
    wishMe()
    while True:

        query = takeCommand()
        if query:
            query = query.lower()
            # proceed with your logic using 'query'
        else:
            speak("Sorry, I didn't catch that.")

        if "friend mode" in query or "friend mod" in query:
            is_friend_mode = True
            chat_history.clear()
            speak("Friend mode activated! Let's talk.")
            continue

        if "assistant mode" in query or "assistant mod" in query:
            if is_friend_mode:
                is_friend_mode = False
                speak("Back to assistant mode.")
            continue

        # ── Friend‑mode chat  (run only while flag is True) ──
        if is_friend_mode:
            if query in ["none", ""]:
                continue                         # ignore silence
            try:
                reply = chat_friend.get_reply(chat_history, query)
            except Exception as e:
                reply = f"Sorry, chat error: {e}"
            chat_history.append({"role": "assistant", "content": reply})
            print("Leo: \n",reply)
            speak(reply)
            continue                             # skip normal command blocks
        
        if "who are you" in query or "what is your name" in query or "tell me about yourself" in query:
            speak("I am Leo, your personal desktop assistant. I can help you with tasks like opening apps, searching the web, playing music, taking screenshots, and much more. Just tell me what to do!")

        elif "female voice" in query:
            engine.setProperty('voice', voices[1].id)
            speak("female voice activated...")

        elif "male voice" in query:
            engine.setProperty('voice', voices[0].id)
            speak("male voice activated...")
        
        elif 'search on wikipedia' in query:  #wikipedia "Your topic" ------------------------------
            query = query.replace("wikipedia", "").strip()
            if not query:
                speak("What should I search on Wikipedia?")
                query = takeCommand().lower()
            try:
                speak('Searching Wikipedia...')
                results = wikipedia.summary(query, sentences=2)
                speak("According to Wikipedia")
                print(results)
                speak(results)
            except wikipedia.exceptions.DisambiguationError as e:
                speak("The topic you gave is too broad. Please be more specific.")
                print("DisambiguationError options:", e.options)
            except wikipedia.exceptions.PageError:
                speak("Sorry, I couldn't find any matching page on Wikipedia.")
            except Exception as e:
                speak("Something went wrong while searching.")
                print("Error:", str(e))
            
        elif 'search on youtube' in query: #search on youtube "Your query" ------------------
            query = query.replace("search on youtube", "").strip()
            webbrowser.get('chrome').open(f"https://www.youtube.com/results?search_query={query}")
            speak(f"Here are the results for {query} on YouTube")

        # Open YouTube and play something #working------------------
        elif 'open youtube' in query:
            speak("What would you like to watch?")
            qrry = takeCommand().lower() #after opening youtube say a command
            kit.playonyt(qrry)
            speak(f"Playing {qrry} on YouTube")

        # Close only YouTube tab in Chrome
        elif 'close youtube' in query: #working ----------------
            closed = False
            for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
                try:
                    if proc.info['name'] == "chrome.exe" and any('youtube' in arg.lower() for arg in proc.info['cmdline']):
                        proc.kill()
                        closed = True
                except (psutil.NoSuchProcess, psutil.AccessDenied, TypeError):
                    continue
            if closed:
                speak("YouTube has been closed.")
            else:
                speak("YouTube is not open in Chrome.")
        
        elif 'close tab' in query and query.strip().lower() == 'close tab':
            try:
                pyautogui.hotkey('ctrl', 'w')
                speak("Closed the current tab.")
            except Exception as e:
                speak("I couldn't close the tab.")
                print(f"[Error - close tab]: {e}")
            
        elif "reopen tab" in query or "reopen closed tab" in query: 
            speak("Reopening the last closed tab.")
            pyautogui.hotkey("ctrl", "shift", "t")


        elif 'search on google' in query:  # search on google "Your query"
            query = query.replace("search on google", "").strip()
            webbrowser.get('chrome').open(f"https://www.google.com/search?q={query}")
            speak(f"Here are the Google search results for {query}")

        # Open Google and search
        elif 'open google' in query: #working ----------------------------------------
            speak("What should I search on Google?")
            qry = takeCommand().lower()  #say your query for searching 
            webbrowser.get('chrome').open(f"https://www.google.com/search?q={qry}")
            try:
                results = wikipedia.summary(qry, sentences=2)
                speak("According to Wikipedia...")
                speak(results)
            except:
                speak("Sorry, I couldn't find anything on Wikipedia.")

        elif "previous tab" in query: #working-----------------------------------------
            pyautogui.hotkey('ctrl', 'shift', 'tab')
            speak("Switched to previous tab.")

        elif "next tab" in query: #working--------------------------------------------
            pyautogui.hotkey('ctrl', 'tab')
            speak("Switched to next tab.")
    
        # Close only Google tab in Chrome
        elif 'close google' in query: #working ---------------------------------
            closed = False
            for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
                try:
                    if proc.info['name'] == "chrome.exe" and any('google' in arg.lower() for arg in proc.info['cmdline']):
                        proc.kill()
                        closed = True
                except (psutil.NoSuchProcess, psutil.AccessDenied, TypeError):
                    continue
            if closed:
                speak("Google tab has been closed.")
            else:
                speak("Google is not open in Chrome.")
            
        elif 'play music' in query: #working------------------------------------------
            music_dir = r"E:\Musics"
            vlc_path = r'"C:\Program Files\VideoLAN\VLC\vlc.exe"'  # Adjust if needed

            try:
                songs = [f for f in os.listdir(music_dir) if f.lower().endswith(('.mp3', '.wav', '.m4a', '.flac'))]
                if songs:
                    playlist = ' '.join([f'"{os.path.join(music_dir, s)}"' for s in songs])
                    os.system(f'start "" {vlc_path} {playlist}')
                    speak(f"Playing your playlist in VLC.")
                else:
                    speak("No songs found.")
            except Exception as e:
                speak("Something went wrong.")
                print(f"[VLC error]: {e}")

        elif 'pause music' in query or 'resume music' in query or 'play or pause music' in query: #working----------
            keyboard.send('play/pause media')
            speak("Music paused or resumed.")
        
        elif 'pause music' in query or 'resume music' in query: #working----------
            press_key(VK_MEDIA_PLAY_PAUSE)
            speak("Toggled play and pause.")

        elif 'next song' in query or 'next track' in query:#working----------
            press_key(VK_MEDIA_NEXT_TRACK)
            speak("Playing next song.")

        elif 'previous song' in query or 'previous track' in query: #working-----------
            press_key(VK_MEDIA_PREV_TRACK)
            speak("Going to previous song.")

        # elif 'stop music' in query:
        #     press_key(VK_MEDIA_STOP)
        #     speak("Music stopped.")

        # elif 'mute' in query:
        #     press_key(VK_VOLUME_MUTE)
        #     speak("Volume muted.")

        elif 'increase volume' in query: #working--------------------------
            for _ in range(5):
                press_key(VK_VOLUME_UP)
            speak("Volume increased.")

        elif 'decrease volume' in query: #working----------------------------
            for _ in range(5):
                press_key(VK_VOLUME_DOWN)
            speak("Volume decreased.")

        elif 'close music' in query: #working--------------------------------------
            # Keywords for music players
            music_keywords = ['vlc', 'wmplayer', 'groove', 'music.ui', 'applicationframehost', 'media.player']
            closed_any = False

            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    pname = proc.info['name'].lower() if proc.info['name'] else ''
                    for keyword in music_keywords:
                        if keyword in pname:
                            proc.kill()
                            closed_any = True
                            # print(f"Killed: {pname}")  # Debug
                            break
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    continue

            if closed_any:
                speak("Music player closed.")
            else:
                speak("No known music player is currently running.")

            
        # elif 'play iron man movie' in query:
        #     npath = "E:\ironman.mkv"
        #     os.startfile(npath)
            
        # elif 'close movie' in query:
        #     os.system("taskkill /f /im vlc.exe")
            
        elif 'tell me the time' in query or 'what is the time' in query: #working----------------------
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")
            
        elif "open notepad" in query: #working but sensitive---------
            subprocess.Popen("notepad.exe")  # open Notepad
            time.sleep(2)  # wait for Notepad to launch

            speak("What should I write in Notepad?")
            time.sleep(1)
            content = takeCommand()
            
            if content and content.lower() != "none":
                pyautogui.typewrite(content)
                time.sleep(1)
                speak("Should I save the file?")
                time.sleep(1)
                answer = takeCommand()

                while answer is None or answer.lower() == "none":
                    speak("I didn't catch that. Should I save the file?")
                    answer = takeCommand()

                if answer and ("yes" in answer.lower() or "save" in answer.lower()):
                    speak("What should be the file name?")
                    time.sleep(1)
                    filename = takeCommand()

                    if filename and filename.lower() != "none":
                        filename = filename.strip().replace(" ", "_")
                        desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
                        full_path = os.path.join(desktop_path, f"{filename}.txt")

                        pyautogui.hotkey('ctrl', 's')  # open Save dialog
                        time.sleep(2)
                        pyautogui.typewrite(full_path)
                        time.sleep(1)
                        pyautogui.press('enter')
                        speak(f"File saved as {filename}.txt on Desktop.")
                    else:
                        speak("I didn't catch the file name. Not saving.")
                else:
                    speak("Okay, not saving the file.")
            else:
                speak("I didn't catch what to write.")

        # === CLOSE NOTEPAD FEATURE ===
        elif "close notepad" in query:#working--------------
            speak("Closing Notepad.")
            os.system("taskkill /f /im notepad.exe")

        elif "open command prompt" in query or "open cmd" in query: #working--------------
            os.system("start cmd")
            
        elif "close command prompt" in query or "close cmd" in query: #working------------
            os.system("taskkill /f /im cmd.exe")

        elif "shutdown the system" in query: #working---------------
            speak("Are you sure you want to shut down the system? Say confirm or yes")
            confirmation = takeCommand()

            if confirmation and ("yes" in confirmation.lower() or "confirm" in confirmation.lower()):
                speak("Shutting down the system in 30 seconds.")
                os.system("shutdown /s /t 30")
            else:
                speak("Shutdown cancelled.")

        elif "cancel shutdown" in query or "stop shutdown" in query or "shutdown cancel" in query: #working--------------
            os.system("shutdown /a")
            speak("Shutdown cancelled.")
        
        elif "shutdown in one minute" in query: #working---------------
            os.system("shutdown /s /t 60")
            speak("System will shut down in one minute.")
            
        elif "restart the system" in query: #working---------------
            os.system("shutdown /r /t 5")

        elif "go to sleep" in query: #working------------------------------
            speak("Going into sleep mode. Say 'wake up' when you need me.")
            while True:
                command = takeCommand()
                if command is None or command.lower() == "none":
                    continue
                if "wake up" in command.lower():
                    speak("I am back online.")
                    break

        elif "full exit" in query: #working------------------------------------
            speak("Are you sure you want me to shut down completely?")
            response = takeCommand()
            if response and ("yes" in response.lower() or "confirm" in response.lower()):
                speak("Alright then, I am shutting down.")
                sys.exit()
            else:
                speak("Shutdown cancelled. I'm still here.")

            
        elif "open camera" in query: #working---------------------------------
            speak("Opening Windows Camera.")
            os.system("start microsoft.windows.camera:")

        elif "click my picture" in query: #working----------------------------
            speak("Clicking your picture.")
            pyautogui.press('enter')
            time.sleep(2)

        elif "save this picture" in query: #working-------------------------------------
            pictures_folder = os.path.join(os.environ["USERPROFILE"], "Pictures", "Camera Roll")
            desktop_folder = os.path.join(os.environ["USERPROFILE"], "Desktop")

            list_of_files = glob.glob(os.path.join(pictures_folder, "*.jpg"))
            if not list_of_files:
                speak("No picture found in Camera Roll.")
            else:
                latest_file = max(list_of_files, key=os.path.getctime)
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                new_filename = os.path.join(desktop_folder, f"Jarvis_Pic_{timestamp}.jpg")
                shutil.copy(latest_file, new_filename)
                speak(f"Picture saved on Desktop as {new_filename.split(os.sep)[-1]}")

        elif "close camera" in query: #working----------------------------------
            speak("Closing the camera.")
            os.system("taskkill /f /im WindowsCamera.exe")

        

            
        elif "take screenshot" in query: #working-------------------------------------
            speak("What should I name the screenshot?")
            name = takeCommand().lower()

            if name == "none" or name.strip() == "":
                speak("I didn't get the name. Screenshot not saved.")
            else:
                speak("Taking screenshot in 3 seconds...")
                time.sleep(3)
                img = pyautogui.screenshot()

                # Build full desktop path
                desktop_path = os.path.join(os.path.join(os.environ["USERPROFILE"]), "Desktop")
                full_path = os.path.join(desktop_path, f"{name}.png")

                img.save(full_path)
                speak(f"Screenshot saved as {name}.png on your desktop.")
            
        elif "calculate" in query: #working------------------------------------
            import operator
            import math

            speak("I'm ready. Please say the expression.")
            r = sr.Recognizer()
            with sr.Microphone() as source:
                r.adjust_for_ambient_noise(source)
                print("Listening for expression...")
                audio = r.listen(source)

            try:
                my_string = r.recognize_google(audio)
                print(f"Recognized: {my_string}")

                expression = my_string.lower()

                # Handle advanced operations first
                if "square root of" in expression:
                    number = int(expression.split("square root of")[1].strip())
                    result = math.sqrt(number)
                elif "factorial of" in expression:
                    number = int(expression.split("factorial of")[1].strip())
                    result = math.factorial(number)
                elif "percent of" in expression or "% of" in expression:
                    if "percent of" in expression:
                        parts = expression.split("percent of")
                    else:
                        parts = expression.split("% of")

                    try:
                        percent = float(parts[0].strip())
                        of_value = float(parts[1].strip())
                        result = (percent / 100) * of_value
                    except:
                        result = None
                elif "to the power" in expression or "raise to" in expression:
                    if "to the power" in expression:
                        parts = expression.split("to the power")
                    else:
                        parts = expression.split("raise to")

                    try:
                        base = float(parts[0].strip())
                        exponent = float(parts[1].strip())
                        result = base ** exponent
                    except:
                        result = None
                else:
                    # Basic math conversion
                    expression = expression.replace("into", "*")
                    expression = expression.replace("multiply", "*")
                    expression = expression.replace("times", "*")
                    expression = expression.replace("x", "*")
                    expression = expression.replace("plus", "+")
                    expression = expression.replace("minus", "-")
                    expression = expression.replace("divided by", "/")
                    expression = expression.replace("divided", "/")
                    result = eval(expression)
                

                print(result)
                speak(f"The result is {result}")
            except Exception as e:
                print(f"Error: {e}")
                speak("Sorry, I couldn't understand or calculate that.")
            
        elif "what is my ip address" in query: #working---------------------------------
            speak("Checking")
            try:
                ipAdd = requests.get('https://api.ipify.org').text
                print(ipAdd)
                speak("your ip adress is")
                speak(ipAdd)
            except Exception as e:
                speak("network is weak, please try again some time later")
                
        elif "volume up" in query: #working--------------------------------------------
            for _ in range(5):  # Increase by ~20-40%
                pyautogui.press("volumeup")
            speak("Volume increased.")
        
        elif "maximum volume" in query: #working-----------------------------------
            for _ in range(50):  # Roughly reaches 100%
                pyautogui.press("volumeup")
            speak("Volume set to maximum.")
            
        elif "volume down" in query: #working-------------------------------------
            for _ in range(5):
                pyautogui.press("volumedown")
            speak("Volume decreased.")
            
        # elif "mute" in query:
        #     pyautogui.press("volumemute")
            
        elif "refresh" in query: #working but sensitive-----------------------------------------------------
            # Step 1: Refresh Desktop
            pyautogui.hotkey("win", "d")  # Show desktop
            time.sleep(1)
            pyautogui.press("f5")        # Refresh desktop
            speak("Desktop refreshed.")
            time.sleep(1)

            # Step 2: Try to activate VS Code from taskbar (simulate Win + 1, Win + 2, etc.)
            # Adjust based on where VS Code is pinned in your taskbar
            try:
                pyautogui.hotkey("win", "9")  # If VS Code is 2nd pinned app in taskbar
                speak("VS Code window brought to front.")
            except Exception as e:
                speak("Couldn't bring VS Code. Trying to open again.")
                subprocess.Popen(r"C:\Users\kumha\AppData\Local\Programs\Microsoft VS Code\Code.exe")
            
        elif "scroll down" in query: #working---------------------------
            pyautogui.moveTo(960, 500)
            pyautogui.scroll(-1000)
            speak("Scrolling down")

        elif "scroll up" in query: #working-------------------------------
            pyautogui.moveTo(960, 500)
            pyautogui.scroll(1000)
            speak("Scrolling up")
        
        elif "minimize the window" in query or "minimise the window" in query: #working-----------------------
            try:
                minimize_active_window()
                speak("Window minimized.")
            except Exception as e:
                speak("Sorry, I couldn't minimize the window.")
                print(e)

        elif "maximize the window" in query: #working--------------------------
            try:
                maximize_last_window()
                speak("Window restored.")
            except Exception as e:
                speak("Sorry, I couldn't restore the window.")
                print(e)

        elif query not in ["", "none", "None"]:  # Fallback handler
            speak("Sorry, that command isn't in my system.")
            with open("unknown_commands.log", "a", encoding="utf-8") as f:
                f.write(f"{datetime.datetime.now()}: {query}\n")
            
        # elif "open paint" in query:
        #     os.system("start mspaint")
        #     speak("Opening Paint")
        #     time.sleep(3)  # Give time for Paint to open

        # elif "select red color" in query or "select red colour":
        #     pyautogui.click(x=180, y=85)  # Update these coordinates
        #     speak("Red color selected.")

        # elif "draw a line" in query:
        #     pyautogui.moveTo(300, 300)  # Starting point
        #     pyautogui.dragTo(500, 300, duration=0.5)  # Ending point
        #     speak("Line drawn.")

        # elif "erase it" in query:
        #     pyautogui.click(x=100, y=85)  # Eraser tool
        #     pyautogui.moveTo(300, 300)
        #     pyautogui.dragTo(500, 300, duration=0.5)
        #     speak("Erased.")

        # elif "close paint" in query:
        #     os.system("taskkill /f /im mspaint.exe")
        #     speak("Paint closed.")