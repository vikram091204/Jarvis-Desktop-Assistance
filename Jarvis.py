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
from functions import settings
import ctypes


class Krishna:
    
    def __init__(self, mode="Microphone", speaker=True):
        self.recognizer = sr.Recognizer()
        self.microphone = sr.Microphone()
        self.engine = pyttsx3.init()
        self.power = None
        self.mode = mode
        self.speaker = speaker
        self.assistant_name = "Krishna"
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
            "find, search file, search folder, locate": self.find_and_open,
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
        
    def listen_for_wake_word(self):
        """Listen for wake word 'Hi Krishna' or similar phrases"""
        if self.mode == "Input":
            user_input = input("Say wake word (Hi Krishna): ").lower()
            wake_detected, command = self.check_wake_word(user_input)
            return wake_detected, command
        
        with self.microphone as source:
            self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
            print("Waiting for wake word 'Hi Krishna'...")
            try:
                audio = self.recognizer.listen(source, timeout=None, phrase_time_limit=5)
                query = self.recognizer.recognize_google(audio).lower()
                print(f"Heard: {query}")
                wake_detected, command = self.check_wake_word(query)
                return wake_detected, command
            except sr.UnknownValueError:
                return False, None
            except sr.RequestError:
                return False, None
            except Exception:
                return False, None
    
    def check_wake_word(self, query):
        """Check if query contains wake word and extract any command after it"""
        wake_words = ["krishna", "hi krishna", "hey krishna", "hello krishna", "ok krishna"]
        
        for wake_word in wake_words:
            if wake_word in query:
                # Extract command after wake word
                parts = query.split(wake_word, 1)
                if len(parts) > 1 and parts[1].strip():
                    command = parts[1].strip()
                    print(f"[DEBUG] Wake word detected with command: '{command}'")
                    return True, command
                return True, None
            
            # Fuzzy matching for variations
            ratio = SequenceMatcher(None, query, wake_word).ratio()
            if ratio > 0.7:
                return True, None
        
        return False, None

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
        best_cmd_length = 0

        for cmd, action in self.commands.items():
            # Split commands and check each variant
            cmd_variants = [c.strip() for c in cmd.split(',')]
            
            for variant in cmd_variants:
                # Check for exact substring match first (highest priority)
                if variant in query or query in variant:
                    ratio = SequenceMatcher(None, query, variant).ratio()
                    # Prioritize longer matches (more specific commands)
                    variant_length = len(variant)
                    
                    # Update if better ratio, or same ratio but longer command
                    if ratio > best_ratio or (ratio == best_ratio and variant_length > best_cmd_length):
                        best_ratio = ratio
                        best_match = action
                        best_cmd = cmd
                        best_cmd_length = variant_length
                else:
                    # Fuzzy matching for non-exact matches
                    ratio = SequenceMatcher(None, query, variant).ratio()
                    if ratio > best_ratio:
                        best_ratio = ratio
                        best_match = action
                        best_cmd = cmd
                        best_cmd_length = len(variant)

        if best_ratio > 0.5 and best_match:
            # Debug output
            print(f"[DEBUG] Matched command: '{best_cmd}' with ratio: {best_ratio:.2f}")
            # Try calling the action with the original query first. Many
            # existing actions have no parameters, so catch TypeError and call
            # without arguments in that case.
            try:
                return best_match(query)
            except TypeError:
                return best_match()

        print(f"[DEBUG] No match found for query: '{query}' (best ratio: {best_ratio:.2f})")
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
        
        print(f"[DEBUG] open_app called with: '{app_query}' -> processed as: '{app}'")

        # Strip common leading verbs such as 'open'
        for prefix in ("open ", "please open ", "could you open "):
            if app.startswith(prefix):
                app = app[len(prefix):].strip()
                print(f"[DEBUG] After stripping prefix: '{app}'")

        # Platform/Website detection - opens in browser if app not installed
        platforms = {
            # Social Media
            'youtube': 'https://www.youtube.com',
            'facebook': 'https://www.facebook.com',
            'instagram': 'https://www.instagram.com',
            'twitter': 'https://www.twitter.com',
            'x': 'https://www.x.com',
            'linkedin': 'https://www.linkedin.com',
            'reddit': 'https://www.reddit.com',
            'tiktok': 'https://www.tiktok.com',
            'snapchat': 'https://www.snapchat.com',
            'pinterest': 'https://www.pinterest.com',
            'whatsapp': 'https://web.whatsapp.com',
            'telegram': 'https://web.telegram.org',
            'discord': 'https://discord.com/app',
            
            # Entertainment
            'netflix': 'https://www.netflix.com',
            'prime video': 'https://www.primevideo.com',
            'amazon prime': 'https://www.primevideo.com',
            'spotify': 'https://open.spotify.com',
            'twitch': 'https://www.twitch.tv',
            
            # Professional/Productivity
            'gmail': 'https://mail.google.com',
            'google drive': 'https://drive.google.com',
            'drive': 'https://drive.google.com',
            'google docs': 'https://docs.google.com',
            'google sheets': 'https://sheets.google.com',
            'google slides': 'https://slides.google.com',
            'outlook': 'https://outlook.live.com',
            'onedrive': 'https://onedrive.live.com',
            'dropbox': 'https://www.dropbox.com',
            
            # Development
            'github': 'https://github.com',
            'gitlab': 'https://gitlab.com',
            'stack overflow': 'https://stackoverflow.com',
            'stackoverflow': 'https://stackoverflow.com',
            
            # Shopping
            'amazon': 'https://www.amazon.com',
            'flipkart': 'https://www.flipkart.com',
            'ebay': 'https://www.ebay.com',
            
            # Other
            'google': 'https://www.google.com',
            'maps': 'https://maps.google.com',
            'google maps': 'https://maps.google.com',
            'translate': 'https://translate.google.com',
            'google translate': 'https://translate.google.com',
        }

        # Check if user wants to open a platform/website
        print(f"[DEBUG] Checking platforms for: '{app}'")
        for platform, url in platforms.items():
            # Check exact match or substring
            if platform in app or app in platform:
                print(f"[DEBUG] Matched platform: '{platform}'")
                webbrowser.open(url)
                message = f"Opening {platform.title()} in browser."
                self.speak(message)
                return message
            # Fuzzy match for speech recognition errors
            ratio = SequenceMatcher(None, app, platform).ratio()
            if ratio > 0.6:
                print(f"[DEBUG] Fuzzy matched platform: '{platform}' with ratio {ratio:.2f}")
                webbrowser.open(url)
                message = f"Opening {platform.title()} in browser."
                self.speak(message)
                return message

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

    def find_and_open(self, query=None):
        """Find a file or folder by name and open it (or open containing folder and select file).
        Searches common user folders, including OneDrive if available. If multiple
        matches are found, presents up to 5 choices and accepts a voice selection
        (number or spoken name).
        """
        if query:
            q = query
        else:
            self.speak("What file or folder should I find?")
            q = self.listen()

        if not q:
            return "No filename provided."

        # Remove leading verbs
        for prefix in ("find ", "search ", "search for ", "locate "):
            if q.startswith(prefix):
                q = q[len(prefix):].strip()

        target = q.lower()

        user_home = os.path.expanduser('~')
        # detect OneDrive paths
        onedrive_candidates = [
            os.environ.get('OneDrive'),
            os.environ.get('OneDriveCommercial'),
            os.environ.get('OneDriveConsumer'),
            os.path.join(user_home, 'OneDrive')
        ]
        search_dirs = [
            os.path.join(user_home, 'Desktop'),
            os.path.join(user_home, 'Documents'),
            os.path.join(user_home, 'Downloads'),
            os.path.join(user_home, 'Pictures'),
            os.path.join(user_home, 'Videos'),
            user_home,
            os.path.join(os.getcwd(), 'data', 'files')
        ]
        for p in onedrive_candidates:
            if p and os.path.exists(p) and p not in search_dirs:
                search_dirs.insert(0, p)

        # Collect matches (path, score)
        matches = []
        best = (None, 0.0)

        for base in search_dirs:
            if not os.path.exists(base):
                continue
            for root, dirs, files in os.walk(base):
                for d in dirs:
                    name = d.lower()
                    path = os.path.join(root, d)
                    if target in name:
                        matches.append((path, 1.0))
                    else:
                        score = SequenceMatcher(None, name, target).ratio()
                        if score > best[1]:
                            best = (path, score)
                for f in files:
                    name = f.lower()
                    path = os.path.join(root, f)
                    if target in name:
                        matches.append((path, 1.0))
                    else:
                        score = SequenceMatcher(None, name, target).ratio()
                        if score > best[1]:
                            best = (path, score)

        # If no direct substring matches, but a good fuzzy match exists, include it
        if not matches and best[0] and best[1] >= 0.6:
            matches.append(best)

        if not matches:
            msg = f"I couldn't find anything matching {q}."
            self.speak(msg)
            return msg

        # Limit to top 5 matches, sorted by score descending
        matches = sorted(matches, key=lambda x: x[1], reverse=True)[:5]

        # If multiple matches, ask user to choose
        if len(matches) > 1:
            speak_list = []
            for idx, (path, score) in enumerate(matches, start=1):
                kind = 'folder' if os.path.isdir(path) else 'file'
                name = os.path.basename(path)
                speak_list.append(f"{idx}. {name} ({kind})")
            self.speak("I found multiple matches:")
            for item in speak_list:
                self.speak(item)
            self.speak("Which one should I open? Say the number or speak the name.")
            choice = self.listen()
            if not choice:
                return "No selection made."

            # try to parse a number
            import re
            m = re.search(r"\d+", choice)
            selected_path = None
            if m:
                n = int(m.group(0))
                if 1 <= n <= len(matches):
                    selected_path = matches[n-1][0]
            if not selected_path:
                # try fuzzy match against basenames
                best_choice = (None, 0.0)
                for path, score in matches:
                    name = os.path.basename(path).lower()
                    s = SequenceMatcher(None, choice, name).ratio()
                    if s > best_choice[1]:
                        best_choice = (path, s)
                if best_choice[1] >= 0.5:
                    selected_path = best_choice[0]
                else:
                    self.speak("Sorry, I couldn't understand your selection.")
                    return "Selection not understood."
        else:
            selected_path = matches[0][0]

        # Open folder or file (select file in explorer)
        try:
            if os.path.isdir(selected_path):
                os.startfile(selected_path)
                msg = f"Opened folder {os.path.basename(selected_path)}."
            else:
                try:
                    subprocess.Popen(["explorer", f"/select,{selected_path}"])
                    msg = f"Opened containing folder and selected {os.path.basename(selected_path)}."
                except Exception:
                    os.startfile(selected_path)
                    msg = f"Opened file {os.path.basename(selected_path)}."
        except Exception as e:
            msg = f"Found item but failed to open: {e}"

        self.speak(msg)
        return msg

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
    krishna = Krishna(mode="Microphone", speaker=True)
    print(f"Krishna Assistant initialized. Say 'Hi Krishna' to activate.")
    try:
        while True:
            # Wait for wake word
            if krishna.listen_for_wake_word():
                krishna.speak("Yes, I'm listening!")
                # Listen for command
                query = krishna.listen()
                if query:
                    krishna.process_query(query)
    except KeyboardInterrupt:
        print("Exiting program...")
        sys.exit(0)


if __name__ == "__main__":
    main()
