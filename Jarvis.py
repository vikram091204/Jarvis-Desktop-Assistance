import datetime
import time
import webbrowser
import sys
import os
import subprocess
import shutil
import requests
from bs4 import BeautifulSoup
import threading
import speech_recognition as sr
from difflib import SequenceMatcher
import pyttsx3
from functions import coder, settings, progress
import ctypes


class Jarvis:
    
    def __init__(self, mode="Microphone", speaker=True):
        self.recognizer = sr.Recognizer()
        self.microphone = sr.Microphone()
        self.engine = pyttsx3.init()
        self.power = None
        self.mode = mode
        self.speaker = speaker
        self.user_name = "Abhay"
        self.commands = {
            "hi, hello, sus": self.greet,
            "what's the time, time right now, what is the time, current time, time please": self.tell_time,
            "what's the date, date today, what is the date, current date, date please": self.tell_date,
            "how are you": self.how_are_you,
            "news": self.get_news,
            "empty recycle bin": self.empty_recycle_bin,
            "restore recycle bin": self.restore_recycle_bin,
            "where is": self.where_is,
            "open camera, click photo": self.open_camera,
            "open browser": self.open_browser,
            "open": self.open_app,
            "play, play song, play video, play music": self.play_media,
            "sleep, lock window": self.lock_window,
            "shutdown": self.shutdown,
            "restart": self.restart,
            "hibernate": self.hibernate,
            "log sign off": self.log_out,
            "cancel": self.cancel,
            "exit, close": self.exit
        }

    def listen(self):
        
        if self.mode == "Input":
            return input("Enter command: ").lower()
        
        with self.microphone as source:
            self.recognizer.adjust_for_ambient_noise(source)
            print("Listening...")
            audio = self.recognizer.listen(source)
        try:
            query = self.recognizer.recognize_google(audio)
            print(f"User said: {query}")
            return query.lower()
        
        except sr.UnknownValueError:
            self.speak("Sorry, I did not understand that.")
            return ""
        except sr.RequestError:
            self.speak("Sorry, my speech service is down.")
            return ""
        
    def speak(self, text):
        print(text)
        
        if self.speaker:
            self.engine.say(text)
            self.engine.runAndWait()

    def process_query(self, query):
        """
        Process user query and return response
        """
        # Process the query and return response (speak returned text)
        try:
            response = self.generate_response(query)
            if isinstance(response, str) and response:
                # Speak the response so voice mode replies audibly
                self.speak(response)
            return response
        except Exception as e:
            err = f"Sorry, I encountered an error: {str(e)}"
            self.speak(err)
            return err
    
    def generate_response(self, query):
        # Find best matching command and attempt to call the handler with the
        # original query when appropriate. Falls back to calling without args.
        best_match = None
        best_ratio = 0
        best_cmd = None

        for cmd, action in self.commands.items():
            ratio = self.similar(query, cmd.split(','))
            if ratio > best_ratio:
                best_ratio = ratio
                best_match = action
                best_cmd = cmd

        if best_ratio > 0.6 and best_match:
            # Try calling the action with the original query first. Many
            # existing actions have no parameters, so catch TypeError and call
            # without arguments in that case.
            try:
                return best_match(query)
            except TypeError:
                return best_match()

        return "Sorry, I didn't understand what you meant."
    
    def similar(self, query, commands):
        
        matchRatioLis = []
        for command in commands:
            matchRatioLis.append(SequenceMatcher(None, query, command).ratio())

            
        return max(matchRatioLis)
    
    def greet(self):
        message = f"Hi {self.user_name} how may I assist you today?"
        self.speak(message)
        return message

    def tell_time(self):
        now = datetime.datetime.now().strftime("%H:%M:%S")
        message = f"The current time is {now}"
        self.speak(message)
        return message

    def tell_date(self):
        today = datetime.datetime.now().strftime("%A, %B %d, %Y")
        message = f"Today is {today}"
        self.speak(message)
        return message

    def how_are_you(self):
        message = "I am fine, thank you. How can I assist you today?"
        self.speak(message)
        return message

    def get_news(self):
        message = "Fetching the latest news for you."
        self.speak(message)
        webbrowser.open("https://news.google.com/")
        return message

    def empty_recycle_bin(self):
        message = "Emptying the recycle bin."
        self.speak(message)
        settings.recycled("empty")
        return message

    def restore_recycle_bin(self):
        message = "Restoring the recycle bin."
        self.speak(message)
        settings.recycled("restore")
        return message

    def where_is(self):
        message = "Please specify the location."
        self.speak(message)
        location = self.listen()
        webbrowser.open(f"https://www.google.com/maps/place/{location}")
        return f"Opening maps for location: {location}"

    def open_camera(self):
        message = "Opening camera."
        self.speak(message)
        settings.camera()
        return message

    def open_browser(self):
        message = "Opening browser."
        self.speak(message)
        webbrowser.open("https://www.google.com")
        return message

    def open_app(self, query=None):
        # If a query is provided (e.g. "open calculator"), use it. Otherwise ask follow-up.
        if query:
            app_query = query
        else:
            message = "Which application would you like to open?"
            self.speak(message)
            app_query = self.listen()

        if not app_query:
            return "No application specified."
        app = app_query.lower()

        # Strip common leading verbs such as 'open'
        for prefix in ("open ", "please open ", "could you open "):
            if app.startswith(prefix):
                app = app[len(prefix):].strip()

        mapping = {
            'file manager': 'explorer',
            'file explorer': 'explorer',
            'explorer': 'explorer',
            'calculator': 'calc.exe',
            'calc': 'calc.exe',
            'notepad': 'notepad.exe',
            'paint': 'mspaint.exe',
            'task manager': 'taskmgr.exe',
            'taskmgr': 'taskmgr.exe',
            'control panel': 'control.exe',
            'settings': 'ms-settings:',
            'command prompt': 'cmd.exe',
            'terminal': 'powershell.exe',
            'powershell': 'powershell.exe',
            'camera': None,
            'browser': 'browser'
        }

        for key, exe in mapping.items():
            if key in app:
                if key == 'camera':
                    return self.open_camera()
                if exe == 'browser':
                    return self.open_browser()
                if isinstance(exe, str) and exe.endswith(':'):
                    try:
                        os.startfile(exe)
                        message = f"Opening {key}."
                    except Exception as e:
                        message = f"Failed to open {key}: {e}"
                    self.speak(message)
                    return message
                try:
                    subprocess.Popen([exe])
                    message = f"Opening {key}."
                except Exception:
                    path = shutil.which(exe)
                    if path:
                        subprocess.Popen([path])
                        message = f"Opening {key}."
                    else:
                        message = f"Could not find executable for {key}."
                self.speak(message)
                return message

        # fallback 1: try to find a file matching the query in common folders
        def find_similar_file(target, search_dirs, threshold=0.6):
            target = target.lower()
            best = (None, 0)
            for base in search_dirs:
                if not os.path.exists(base):
                    continue
                for root, dirs, files in os.walk(base):
                    for f in files:
                        name = f.lower()
                        # direct substring match strong score
                        if target in name:
                            return os.path.join(root, f)
                        # fuzzy match on filename without extension
                        score = SequenceMatcher(None, name, target).ratio()
                        if score > best[1]:
                            best = (os.path.join(root, f), score)
            if best[1] >= threshold:
                return best[0]
            return None

        user_home = os.path.expanduser('~')
        search_dirs = [
            os.path.join(user_home, 'Desktop'),
            os.path.join(user_home, 'Documents'),
            os.path.join(user_home, 'Downloads'),
            os.path.join(os.getcwd(), 'data', 'files')
        ]

        file_match = find_similar_file(app, search_dirs)
        if file_match:
            try:
                os.startfile(file_match)
                message = f"Opening file {os.path.basename(file_match)}."
            except Exception as e:
                message = f"Found file but failed to open it: {e}"
            self.speak(message)
            return message

        # fallback 2: try to open as a path or executable name
        try:
            os.startfile(app_query)
            message = f"Opening {app_query}."
            self.speak(message)
            return message
        except Exception:
            exe = shutil.which(app.split()[0])
            if exe:
                try:
                    subprocess.Popen([exe])
                    message = f"Opening {app_query}."
                    self.speak(message)
                    return message
                except Exception:
                    pass

        message = f"Sorry, I couldn't open {app_query}."
        self.speak(message)
        return message

    def play_media(self, query=None):
        """Play a song or video. Accepts an optional query string (e.g. "play hotel california").
        Searches common media folders for matching filenames and opens the best match,
        or opens URLs (YouTube) in the browser.
        """
        # Determine target from query or ask follow-up
        if query:
            target = query
        else:
            self.speak("Which song or video should I play?")
            target = self.listen()
        if not target:
            return "No media specified."

        target = target.lower()
        # strip leading verbs
        for prefix in ("play ", "please play "):
            if target.startswith(prefix):
                target = target[len(prefix):].strip()

        # Detect if user asked to play on YouTube/online explicitly
        prefer_online = False
        online_indicators = ('on youtube', 'youtube', 'youtu.be', 'online', 'yt')
        for ind in online_indicators:
            if ind in target:
                prefer_online = True
                target = target.replace(ind, '').strip()

        # If it's a URL, play it directly
        if target.startswith('http'):
            self.speak(f"Playing {target} in browser")
            webbrowser.open(target)
            return f"Playing {target}"

        # search common media directories
        user_home = os.path.expanduser('~')
        media_dirs = [
            os.path.join(user_home, 'Music'),
            os.path.join(user_home, 'Videos'),
            os.path.join(user_home, 'Downloads'),
            os.path.join(user_home, 'Desktop'),
            os.path.join(os.getcwd(), 'data', 'files')
        ]

        # file extensions to consider as media
        media_exts = ('.mp3', '.wav', '.flac', '.m4a', '.mp4', '.mkv', '.avi', '.mov')

        best = (None, 0)
        for base in media_dirs:
            if not os.path.exists(base):
                continue
            for root, dirs, files in os.walk(base):
                for f in files:
                    name = f.lower()
                    if not name.endswith(media_exts):
                        continue
                    if target in name:
                        # immediate match
                        best = (os.path.join(root, f), 1.0)
                        break
                    score = SequenceMatcher(None, name, target).ratio()
                    if score > best[1]:
                        best = (os.path.join(root, f), score)
                if best[1] == 1.0:
                    break
            if best[1] == 1.0:
                break

        if best[0] and best[1] >= 0.55 and not prefer_online:
            try:
                os.startfile(best[0])
                message = f"Playing {os.path.basename(best[0])}."
            except Exception as e:
                message = f"Found media but failed to play: {e}"
            self.speak(message)
            return message

        # fallback: try to open any file matching (not just media exts)
        for base in media_dirs:
            if not os.path.exists(base):
                continue
            for root, dirs, files in os.walk(base):
                for f in files:
                    name = f.lower()
                    if target in name:
                        path = os.path.join(root, f)
                        try:
                            os.startfile(path)
                            message = f"Opening {os.path.basename(path)}."
                        except Exception as e:
                            message = f"Found file but failed to open: {e}"
                        self.speak(message)
                        return message

        # If not found, try to play the top YouTube result (online)
        query_str = target.replace(' ', '+')
        yt_search = f"https://www.youtube.com/results?search_query={query_str}"
        self.speak(f"I couldn't find a local file. Searching YouTube for {target}.")
        try:
            resp = requests.get(yt_search, headers={"User-Agent": "Mozilla/5.0"}, timeout=8)
            html = resp.text
            video_url = None
            # First try parsing anchors
            soup = BeautifulSoup(html, 'html.parser')
            for a in soup.find_all('a', href=True):
                href = a['href']
                if href.startswith('/watch'):
                    video_url = 'https://www.youtube.com' + href
                    break
            # Fallback: regex search for watch?v= pattern in page HTML
            if not video_url:
                import re
                m = re.search(r"/watch\?v=[A-Za-z0-9_-]{6,}", html)
                if m:
                    video_url = 'https://www.youtube.com' + m.group(0)
            if video_url:
                webbrowser.open(video_url)
                self.speak(f"Playing {target} on YouTube.")
                return f"Playing {video_url}"
        except Exception:
            pass

        # fallback: open the search results page
        webbrowser.open(yt_search)
        return f"Searched YouTube for {target}"

    def lock_window(self):
        message = "Locking the window."
        self.speak(message)
        self.power = threading.Thread(target=settings.power,args=['lock'], daemon=True)
        self.power.start()
        return message

    def shutdown(self):
        message = "Shutting down the system in 20 seconds.\nPlease close any opened applications."
        self.speak(message)
        self.power = threading.Thread(target=settings.power,args=['shutdown'], daemon=True)
        self.power.start()
        return message

    def restart(self):
        message = "Restarting the system in 20 seconds.\nPlease close any opened applications."
        self.speak(message)
        self.power = threading.Thread(target=settings.power,args=['restart'], daemon=True)
        self.power.start()
        return message

    def hibernate(self):
        message = "Hibernating the system in 20 seconds."
        self.speak(message)
        self.power = threading.Thread(target=settings.power,args=['hibernate'], daemon=True)
        self.power.start()
        return message

    def log_out(self):
        message = "Signing off the system in 10 seconds.\nPlease close any opened applications."
        self.speak(message)
        self.power = threading.Thread(target=settings.power,args=['log out'], daemon=True)
        self.power.start()
        return message

    def cancel(self):
        if self.power is not None and self.power.is_alive():
            self.stopThread(self.power)
            return "Cancelled the pending operation."
        return "No operation to cancel."

    def exit(self):
        message = "Exiting program..."
        self.speak(message)
        time.sleep(1)
        sys.exit(0)
        return message

    def getThreadId(self, thread):

        if hasattr(thread,'_thread_id'):
            return thread._thread_id
        for id,thread1 in threading._active.items():
            if thread1 is thread:
                return id

    def stopThread(self, thread):
        threadId = self.getThreadId(thread)
        res = ctypes.pythonapi.PyThreadState_SetAsyncExc(threadId,\
            ctypes.py_object(SystemExit))

        if res > 1:
            ctypes.pythonapi.PyThreadState_SetAsyncExc(threadId,0)


def main():
    jarvis = Jarvis(mode="Microphone", speaker=True)
    try:
        while True:
            query = jarvis.listen()
            if query:
                jarvis.process_query(query)
    except KeyboardInterrupt:
        print("Exiting program...")
        sys.exit(0)


if __name__ == "__main__":
    main()
