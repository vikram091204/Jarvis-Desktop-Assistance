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
from pathlib import Path


class Siri:
    
    def __init__(self, mode="Microphone", speaker=True, voice_index=None):
        self.recognizer = sr.Recognizer()
        self.microphone = sr.Microphone()
        self.engine = pyttsx3.init()
        self.power = None
        self.mode = mode
        self.speaker = speaker
        self.voice_index = voice_index
        # If a voice index is provided and valid, set the TTS voice
        try:
            if voice_index is not None:
                voices = self.engine.getProperty('voices')
                if 0 <= int(voice_index) < len(voices):
                    self.engine.setProperty('voice', voices[int(voice_index)].id)
                else:
                    print(f"Voice index {voice_index} out of range. Available: 0..{len(voices)-1}")
        except Exception as e:
            print(f"Failed to set voice: {e}")
        self.assistant_name = "Siri"
        self.user_name = "Vikram"
        # Ensure avatar state file exists and (try to) start avatar UI
        try:
            data_dir = Path(__file__).parent.joinpath('data')
            data_dir.mkdir(parents=True, exist_ok=True)
            state_file = data_dir.joinpath('avatar_state.txt')
            if not state_file.exists():
                state_file.write_text('idle')

            # Auto-start avatar.py in a separate process so the floating widget appears
            avatar_path = Path(__file__).parent.joinpath('avatar.py')
            if avatar_path.exists():
                try:
                    # Use the same Python interpreter
                    cmd = [sys.executable, str(avatar_path)]
                    # Start detached process; do not wait
                    if os.name == 'nt':
                        DETACHED = 0x00000008
                        subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=DETACHED)
                    else:
                        subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                except Exception:
                    pass
        except Exception:
            pass
        self.commands = {
            "hi, hello, sus": self.greet,
            "what's the time, time right now, what is the time, current time, time please": self.tell_time,
            "what's the date, date today, what is the date, current date, date please": self.tell_date,
            "how are you": self.how_are_you,
            "news": self.get_news,
            "empty recycle bin": self.empty_recycle_bin,
            "restore recycle bin": self.restore_recycle_bin,
            "where is": self.where_is,
            "open": self.open_app,
            "play, play song, play video, play music": self.play_media,
            "find, search file, search folder, locate": self.find_and_open,
            "list voices, voices": self.list_voices,
            "change voice, set voice, voice": self.change_voice,
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
        
        # indicate avatar that we're listening (if avatar app is running)
        try:
            Path = __import__('pathlib').Path
            state_file = Path(__file__).parent.joinpath('data', 'avatar_state.txt')
            state_file.parent.mkdir(parents=True, exist_ok=True)
            state_file.write_text('listening')
        except Exception:
            pass

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
            try:
                Path = __import__('pathlib').Path
                state_file = Path(__file__).parent.joinpath('data', 'avatar_state.txt')
                state_file.write_text('idle')
            except Exception:
                pass
            return ""
        except sr.RequestError:
            self.speak("Sorry, my speech service is down.")
            try:
                Path = __import__('pathlib').Path
                state_file = Path(__file__).parent.joinpath('data', 'avatar_state.txt')
                state_file.write_text('idle')
            except Exception:
                pass
            return ""
        
    def listen_for_wake_word(self):
        """Listen for wake word 'Hi Siri' or similar phrases"""
        if self.mode == "Input":
            user_input = input("Say wake word (Hi Siri): ").lower()
            wake_detected, command = self.check_wake_word(user_input)
            return wake_detected, command
        
        # indicate listening/waiting for wake word
        try:
            Path = __import__('pathlib').Path
            state_file = Path(__file__).parent.joinpath('data', 'avatar_state.txt')
            state_file.parent.mkdir(parents=True, exist_ok=True)
            state_file.write_text('listening')
        except Exception:
            pass

        with self.microphone as source:
            self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
            print("Waiting for wake word 'Hi Siri'...")
            try:
                audio = self.recognizer.listen(source, timeout=None, phrase_time_limit=5)
                query = self.recognizer.recognize_google(audio).lower()
                print(f"Heard: {query}")
                wake_detected, command = self.check_wake_word(query)
                # once we've processed wake-word, go idle
                try:
                    Path = __import__('pathlib').Path
                    state_file = Path(__file__).parent.joinpath('data', 'avatar_state.txt')
                    state_file.write_text('idle')
                except Exception:
                    pass
                return wake_detected, command
            except sr.UnknownValueError:
                return False, None
            except sr.RequestError:
                return False, None
            except Exception:
                try:
                    Path = __import__('pathlib').Path
                    state_file = Path(__file__).parent.joinpath('data', 'avatar_state.txt')
                    state_file.write_text('idle')
                except Exception:
                    pass
                return False, None
    
    def check_wake_word(self, query):
        """Check if query contains wake word and extract any command after it"""
        wake_words = ["siri", "hi siri", "hey siri", "hello siri", "ok siri"]
        
        for wake_word in wake_words:
            if wake_word in query:
                # Extract command after wake word
                parts = query.split(wake_word, 1)
                if len(parts) > 1 and parts[1].strip():
                    command = parts[1].strip()

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
            # Notify avatar that we're speaking
            try:
                Path = __import__('pathlib').Path
                state_file = Path(__file__).parent.joinpath('data', 'avatar_state.txt')
                state_file.parent.mkdir(parents=True, exist_ok=True)
                state_file.write_text('speaking')
            except Exception:
                pass

            self.engine.say(text)
            self.engine.runAndWait()

            # Back to idle after speaking
            try:
                Path = __import__('pathlib').Path
                state_file = Path(__file__).parent.joinpath('data', 'avatar_state.txt')
                state_file.write_text('idle')
            except Exception:
                pass

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
        best_variant = None
        best_cmd_length = 0

        for cmd, action in self.commands.items():
            # Split commands and check each variant
            cmd_variants = [c.strip() for c in cmd.split(',')]
            
            for variant in cmd_variants:
                # Check for exact substring match first (highest priority)
                if variant in query or query in variant:
                    # When the variant appears as a substring in the query
                    # treat it as a strong match (boost the ratio) so short
                    # variants like "open" match "open browser" reliably.
                    ratio = max(SequenceMatcher(None, query, variant).ratio(), 0.95)
                    # Prioritize longer matches (more specific commands)
                    variant_length = len(variant)

                    # Update if better ratio, or same ratio but longer command
                    if ratio > best_ratio or (ratio == best_ratio and variant_length > best_cmd_length):
                        best_ratio = ratio
                        best_match = action
                        best_cmd = cmd
                        best_variant = variant
                        best_cmd_length = variant_length
                else:
                    # Fuzzy matching for non-exact matches
                    ratio = SequenceMatcher(None, query, variant).ratio()
                    if ratio > best_ratio:
                        best_ratio = ratio
                        best_match = action
                        best_cmd = cmd
                        best_variant = variant
                        best_cmd_length = len(variant)

        # Increased threshold from 0.5 to 0.7 to prevent false matches
        if best_ratio >= 0.7 and best_match:
            # Debug message removed to keep terminal output clean.
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
        return message

    def tell_time(self):
        now = datetime.datetime.now().strftime("%H:%M:%S")
        message = f"The current time is {now}"
        return message

    def tell_date(self):
        today = datetime.datetime.now().strftime("%A, %B %d, %Y")
        message = f"Today is {today}"
        return message

    def list_voices(self):
        """List available TTS voices (prints and speaks brief summary)."""
        try:
            voices = self.engine.getProperty('voices')
            if not voices:
                return "I could not find any voices installed."

            # Provide only two options: male and female. Choose the first
            # available voice that appears to be male and the first female one.
            male = None
            female = None
            for i, v in enumerate(voices):
                gender = getattr(v, 'gender', None)
                name = getattr(v, 'name', '')
                lname = name.lower()
                if not female and ((gender and 'female' in str(gender).lower()) or 'zira' in lname or 'hazel' in lname):
                    female = (i, v)
                if not male and ((gender and 'male' in str(gender).lower()) or 'david' in lname or 'male' in lname):
                    male = (i, v)

            # Fallback: if genders not found, pick first two distinct voices
            if not female and voices:
                female = (0, voices[0])
            if not male and len(voices) > 1:
                male = (1, voices[1])

            # Announce the two options
            if female:
                self.speak(f"Female voice option: {getattr(female[1],'name',female[1].id)}")
                print(f"female: [{female[0]}] {getattr(female[1],'name',female[1].id)}")
            if male:
                self.speak(f"Male voice option: {getattr(male[1],'name',male[1].id)}")
                print(f"male: [{male[0]}] {getattr(male[1],'name',male[1].id)}")

            return "Listed male/female voice options."
        except Exception as e:
            err = f"Failed to list voices: {e}"
            return err

    def change_voice(self, query=None):
        """Change the TTS voice at runtime. Accepts an index, gender, or name.

        Examples:
        - "change voice to 1"
        - "set voice 2"
        - "change voice to female"
        - "change voice to zira"
        """
        try:
            voices = self.engine.getProperty('voices')
            if query:
                q = query.lower()
            else:
                self.speak("Which voice would you like? Say the number or name.")
                q = self.listen()
            if not q:
                return "No voice specified."

            # Only accept 'male' or 'female' options per user's request
            if 'female' in q:
                # find first female-like voice
                for i, v in enumerate(voices):
                    gender = getattr(v, 'gender', None)
                    name = getattr(v, 'name', '').lower()
                    if (gender and 'female' in str(gender).lower()) or 'zira' in name or 'hazel' in name:
                        self.engine.setProperty('voice', v.id)
                        self.voice_index = i
                        msg = f"Voice changed to {getattr(v,'name',v.id)}."
                        return msg
                # fallback
                self.engine.setProperty('voice', voices[0].id)
                self.voice_index = 0
                msg = f"Voice changed to {getattr(voices[0],'name',voices[0].id)}."
                return msg

            if 'male' in q:
                # find first male-like voice
                for i, v in enumerate(voices):
                    gender = getattr(v, 'gender', None)
                    name = getattr(v, 'name', '').lower()
                    if (gender and 'male' in str(gender).lower()) or 'david' in name:
                        self.engine.setProperty('voice', v.id)
                        self.voice_index = i
                        msg = f"Voice changed to {getattr(v,'name',v.id)}."
                        return msg
                # fallback
                if len(voices) > 1:
                    self.engine.setProperty('voice', voices[1].id)
                    self.voice_index = 1
                    msg = f"Voice changed to {getattr(voices[1],'name',voices[1].id)}."
                    return msg

            msg = "Please say 'male' or 'female' to change voice. Say 'list voices' to hear options."
            return msg
        except Exception as e:
            err = f"Failed to change voice: {e}"
            return err

    def how_are_you(self):
        message = "I am fine, thank you. How can I assist you today?"
        return message

    def get_news(self):
        message = "Fetching the latest news for you."
        webbrowser.open("https://news.google.com/")
        return message

    def empty_recycle_bin(self):
        message = "Emptying the recycle bin."
        settings.recycled("empty")
        return message

    def restore_recycle_bin(self):
        message = "Restoring the recycle bin."
        settings.recycled("restore")
        return message

    def where_is(self):
        message = "Please specify the location."
        location = self.listen()
        webbrowser.open(f"https://www.google.com/maps/place/{location}")
        return f"Opening maps for location: {location}"

    def open_camera(self):
        message = "Opening camera."
        settings.camera()
        return message

    def open_browser(self):
        message = "Opening browser."
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
        
        # Strip common leading verbs (English pattern: "open YouTube")
        for prefix in ("open ", "please open ", "could you open "):
            if app.startswith(prefix):
                app = app[len(prefix):].strip()
        
        # Strip common trailing action words (Hindi/Hinglish pattern: "YouTube open kro")
        for suffix in (" open kro", " kholo", " chalu kro", " open karo", " karo", " kro", 
                       " open kar do", " open kar", " chalao", " start kro", " start karo"):
            if app.endswith(suffix):
                app = app[:-len(suffix)].strip()


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
            'browser': 'https://www.google.com',
            'maps': 'https://maps.google.com',
            'google maps': 'https://maps.google.com',
            'translate': 'https://translate.google.com',
            'google translate': 'https://translate.google.com',
        }

        # Check if user wants to open a platform/website. Prefer launching
        # a locally installed application if available; otherwise open the
        # website in the browser.

        # Mapping of platform keys to likely executable names (best-effort).
        exe_mapping = {
            'spotify': ['spotify.exe', 'Spotify.exe'],
            'discord': ['Discord.exe', 'discord.exe'],
            'telegram': ['Telegram.exe', 'telegram.exe'],
            'whatsapp': ['WhatsApp.exe', 'WhatsAppDesktop.exe'],
            'youtube': ['YouTube.exe', 'VLC.exe', 'vlc.exe', 'mpv.exe', 'PotPlayer.exe', 'PotPlayerMini64.exe', 'Kodi.exe'],
            'vscode': ['Code.exe', 'code.exe'],
            'code': ['Code.exe', 'code.exe'],
            'chrome': ['chrome.exe'],
            'firefox': ['firefox.exe'],
            'edge': ['msedge.exe', 'msedgewebview2.exe'],
            'slack': ['slack.exe'],
            'zoom': ['Zoom.exe', 'zoom.exe'],
            'steam': ['steam.exe']
        }

        def find_executable(exe_names):
            """Try to locate an executable by name. Returns full path or None."""
            # 1) try shutil.which
            for name in exe_names:
                path = shutil.which(name)
                if path:
                    return path

            # 2) search common Program Files directories for a matching exe
            program_dirs = [os.environ.get('ProgramFiles'), os.environ.get('ProgramFiles(x86)'), os.environ.get('LOCALAPPDATA')]
            checked = set()
            for base in program_dirs:
                if not base or base in checked:
                    continue
                checked.add(base)
                for root, dirs, files in os.walk(base):
                    for f in files:
                        for candidate in exe_names:
                            if f.lower() == candidate.lower():
                                return os.path.join(root, f)
            return None

        for platform, url in platforms.items():
            # Check exact match or substring
            if platform in app or app in platform:
                # If we have known executables for this platform, try to launch them
                exe_candidates = exe_mapping.get(platform)
                if exe_candidates:
                    exe_path = find_executable(exe_candidates)
                    if exe_path:
                        try:
                            # If this is YouTube, try to pass the site URL to capable players
                            if platform == 'youtube':
                                try:
                                    message = f"Opening {platform.title()}."
                                    subprocess.Popen([exe_path, 'https://www.youtube.com'])
                                    return message
                                except Exception:
                                    # Try launching without args if the app doesn't accept URL
                                    try:
                                        subprocess.Popen([exe_path])
                                        message = f"Opening {platform.title()}."
                                        return message
                                    except Exception:
                                        pass
                            else:
                                message = f"Opening {platform.title()}."
                                subprocess.Popen([exe_path])
                                return message
                        except Exception:
                            # Fall back to opening URL if launching fails
                            pass

                # Special-case: try WhatsApp protocol before opening web client
                if platform == 'whatsapp':
                    for proto in ('whatsapp://', 'whatsapp://send'):
                        try:
                            os.startfile(proto)
                            message = "Opening WhatsApp application."
                            self.speak(message)
                            return message
                        except Exception:
                            pass

                # Fuzzy match or no local app found: open the website
                message = f"Opening {platform.title()}."
                webbrowser.open(url)
                return message

            # Fuzzy match for speech recognition errors
            ratio = SequenceMatcher(None, app, platform).ratio()
            if ratio > 0.6:
                exe_candidates = exe_mapping.get(platform)
                if exe_candidates:
                    exe_path = find_executable(exe_candidates)
                    if exe_path:
                        try:
                            message = f"Opening {platform.title()}."
                            subprocess.Popen([exe_path])
                            return message
                        except Exception:
                            pass
                # Special-case WhatsApp protocol before web fallback
                if platform == 'whatsapp':
                    for proto in ('whatsapp://', 'whatsapp://send'):
                        try:
                            os.startfile(proto)
                            message = "Opening WhatsApp application."
                            self.speak(message)
                            return message
                        except Exception:
                            pass

                message = f"Opening {platform.title()}."
                webbrowser.open(url)
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
            'photo': None,
            'click photo': None,
            'browser': 'browser'
        }

        for key, exe in mapping.items():
            if key in app:
                if key in ['camera', 'photo', 'click photo']:
                    return self.open_camera()
                if exe == 'browser':
                    return self.open_browser()
                if isinstance(exe, str) and exe.endswith(':'):
                    try:
                        os.startfile(exe)
                        message = f"Opening {key}."
                    except Exception as e:
                        message = f"Failed to open {key}: {e}"
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
            return message

        # fallback 2: try to open as a path or executable name
        try:
            os.startfile(app_query)
            message = f"Opening {app_query}."
            return message
        except Exception:
            exe = shutil.which(app.split()[0])
            if exe:
                try:
                    subprocess.Popen([exe])
                    message = f"Opening {app_query}."
                    return message
                except Exception:
                    pass

        message = f"Sorry, I couldn't open {app_query}."
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
        # strip leading verbs (English pattern: "play hotel california")
        for prefix in ("play ", "please play "):
            if target.startswith(prefix):
                target = target[len(prefix):].strip()
        
        # strip trailing action words (Hindi/Hinglish pattern: "hotel california bajao")
        for suffix in (" play kro", " play karo", " bajao", " chalao", " sunao", 
                       " play kar do", " play kar", " laga do", " chala do"):
            if target.endswith(suffix):
                target = target[:-len(suffix)].strip()

        # Detect if user asked to play on YouTube/online explicitly
        prefer_online = False
        online_indicators = ('on youtube', 'youtube', 'youtu.be', 'online', 'yt')
        for ind in online_indicators:
            if ind in target:
                prefer_online = True
                target = target.replace(ind, '').strip()

        # If it's a URL, play it directly
        if target.startswith('http'):
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
                        return message

        # If not found, try to play the top YouTube result (online)
        query_str = target.replace(' ', '+')
        yt_search = f"https://www.youtube.com/results?search_query={query_str}"
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

        # Remove leading verbs (English pattern: "find document")
        for prefix in ("find ", "search ", "search for ", "locate "):
            if q.startswith(prefix):
                q = q[len(prefix):].strip()
        
        # Remove trailing action words (Hindi/Hinglish pattern: "document dhundo")
        for suffix in (" dhundo", " khojo", " find kro", " find karo", " search kro", 
                       " find kar do", " locate kro", " dhoondho"):
            if q.endswith(suffix):
                q = q[:-len(suffix)].strip()

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

        return msg

    def lock_window(self):
        message = "Locking the window."
        self.power = threading.Thread(target=settings.power,args=['lock'], daemon=True)
        self.power.start()
        return message

    def shutdown(self):
        message = "Shutting down the system in 20 seconds.\nPlease close any opened applications."
        self.power = threading.Thread(target=settings.power,args=['shutdown'], daemon=True)
        self.power.start()
        return message

    def restart(self):
        message = "Restarting the system in 20 seconds.\nPlease close any opened applications."
        self.power = threading.Thread(target=settings.power,args=['restart'], daemon=True)
        self.power.start()
        return message

    def hibernate(self):
        message = "Hibernating the system in 20 seconds."
        self.power = threading.Thread(target=settings.power,args=['hibernate'], daemon=True)
        self.power.start()
        return message

    def log_out(self):
        message = "Signing off the system in 10 seconds.\nPlease close any opened applications."
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
    # Default to a female voice (voice_index=2 -> Microsoft Zira on Windows)
    siri = Siri(mode="Microphone", speaker=True, voice_index=2)
    print(f"Siri Assistant initialized. Say 'Hi Siri' to activate.")
    try:
        while True:
            # Wait for wake word
            wake_detected, command_from_wake = siri.listen_for_wake_word()
            if wake_detected:
                siri.speak("Yes, I'm listening!")
                
                # Enter continuous listening mode
                active = True
                while active:
                    # If command was included in wake phrase, use it
                    if command_from_wake:
                        siri.process_query(command_from_wake)
                        command_from_wake = None  # Clear it
                    else:
                        # Listen for command
                        query = siri.listen()
                        if query:
                            # Check for sleep/exit commands
                            if any(word in query for word in ['sleep', 'go to sleep', 'exit', 'goodbye', 'bye']):
                                siri.speak("Going to sleep. Say Hi Siri to wake me up.")
                                active = False
                            else:
                                siri.process_query(query)
    except KeyboardInterrupt:
        print("Exiting program...")
        sys.exit(0)


if __name__ == "__main__":
    main()
