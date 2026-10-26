"""
Jarvis Voice Assistant (Windows)

Install the requirements once:
    pip install SpeechRecognition pyttsx3 pyaudio comtypes pywin32 pywhatkit

musicLibrary.py (optional) can sit in the same folder as this file.
Your songs are also built in below, so music works either way.
"""

import datetime
import difflib
import importlib.util
import os
import re
import socket
import sys
import time
import urllib.parse
import webbrowser
import subprocess

import speech_recognition as sr

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# -------------------- CHROME SETUP --------------------


def find_chrome():
    """Look for chrome.exe in the usual Windows install folders."""
    candidates = [
        os.path.expandvars(r"%PROGRAMFILES%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%PROGRAMFILES(X86)%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    ]
    return next((p for p in candidates if os.path.isfile(p)), None)


CHROME_PATH = find_chrome()

if CHROME_PATH:
    # Makes webbrowser.open() (and pywhatkit) use Chrome instead of Edge
    webbrowser.register(
        "chrome", None, webbrowser.BackgroundBrowser(CHROME_PATH), preferred=True
    )
    print(f"Using Chrome: {CHROME_PATH}")
else:
    print("Google Chrome was not found, using the default browser instead.")


def open_url(url):
    """Open a URL in Chrome, or the default browser if Chrome is missing."""
    if CHROME_PATH:
        subprocess.Popen([CHROME_PATH, url])
    else:
        webbrowser.open(url)


# -------------------- MUSIC LIBRARY --------------------

# Built-in songs, so the library is never empty.
BUILTIN_MUSIC = {
    "tu janna": "https://www.youtube.com/watch?v=1wYNFfgrXTI",
    "deserve": "https://youtu.be/1bIk_K_OZL4?si=czd8LK0s8ZurpnVL",
    "attention": "https://www.youtube.com/watch?v=0Vw3v8c1Z6M",
    "playlist1": "https://www.youtube.com/playlist?list=PLqG35nJ9NfR1X6R1X6R1X6R1X6R1X6R1X",
    "alvida alvida": "https://youtu.be/KaUESRM2nt4?si=zePTBwSxDttpBRRw",
}


def load_music():
    """Built-in songs + anything found in musicLibrary.py next to this file."""
    songs = dict(BUILTIN_MUSIC)
    path = os.path.join(BASE_DIR, "musicLibrary.py")

    if os.path.exists(path):
        try:
            spec = importlib.util.spec_from_file_location("musicLibrary", path)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            extra = {
                str(name).lower().strip(): url
                for name, url in module.music.items()
            }
            songs.update(extra)
            print(f"Loaded musicLibrary.py ({len(extra)} songs).")
        except Exception as error:
            print(f"Could not read musicLibrary.py: {error}")
            print("Using the built-in songs instead.")
    else:
        print(f"musicLibrary.py not found in {BASE_DIR}")
        print("Using the built-in songs instead.")

    return songs


music = load_music()
print("Songs available: " + ", ".join(music))

# -------------------- TEXT-TO-SPEECH SETUP --------------------
# Two speech backends are used so Jarvis can always talk:
#   1. Windows SAPI directly (pywin32)  -> most reliable
#   2. pyttsx3 with a FRESH engine each time -> fallback

try:
    import win32com.client
    import pythoncom
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

try:
    import pyttsx3
    HAS_PYTTSX3 = True
except ImportError:
    HAS_PYTTSX3 = False

# Optional: lets Jarvis start a YouTube video automatically
# for songs that are NOT in your music library.
try:
    import pywhatkit
    HAS_PYWHATKIT = True
except Exception:
    HAS_PYWHATKIT = False

if not HAS_WIN32 and not HAS_PYTTSX3:
    print("No speech engine found. Run: pip install pywin32 pyttsx3")

_sapi_voice = None


def _get_sapi_voice():
    """Create the Windows SAPI voice once and prefer the Zira voice."""
    global _sapi_voice

    if _sapi_voice is None:
        pythoncom.CoInitialize()
        _sapi_voice = win32com.client.Dispatch("SAPI.SpVoice")
        _sapi_voice.Volume = 100
        _sapi_voice.Rate = 1  # range -10 (slow) to 10 (fast)

        try:
            voices = _sapi_voice.GetVoices()
            for i in range(voices.Count):
                if "zira" in voices.Item(i).GetDescription().lower():
                    _sapi_voice.Voice = voices.Item(i)
                    break
        except Exception as error:
            print(f"(Could not select Zira voice: {error})")

    return _sapi_voice


def _speak_sapi(text):
    _get_sapi_voice().Speak(text)


def _speak_pyttsx3(text):
    engine = pyttsx3.init("sapi5")  # fresh engine every time
    engine.setProperty("rate", 175)
    engine.setProperty("volume", 1.0)

    try:
        for voice in engine.getProperty("voices"):
            if "zira" in (voice.name + " " + voice.id).lower():
                engine.setProperty("voice", voice.id)
                break
    except Exception as error:
        print(f"(Could not select Zira voice: {error})")

    engine.say(text)
    engine.runAndWait()
    engine.stop()


def speak(text):
    """Print AND speak the text out loud."""
    print(f"\nJarvis: {text}")

    if HAS_WIN32:
        try:
            _speak_sapi(text)
            return
        except Exception as error:
            print(f"Speech error (SAPI): {error}")

    if HAS_PYTTSX3:
        try:
            _speak_pyttsx3(text)
            return
        except Exception as error:
            print(f"Speech error (pyttsx3): {error}")

    print("(Jarvis could not speak. Check your speakers and installs.)")


# -------------------- SPEECH RECOGNITION SETUP --------------------

recognizer = sr.Recognizer()
recognizer.dynamic_energy_threshold = True
recognizer.pause_threshold = 0.8


def listen():
    """Listen to one sentence and convert it to text."""
    try:
        with sr.Microphone() as source:
            print("\nListening...")
            audio = recognizer.listen(
                source,
                timeout=6,
                phrase_time_limit=8
            )

    except sr.WaitTimeoutError:
        return ""

    except OSError as error:
        print(f"Microphone error: {error}")
        speak("I cannot access your microphone, ma'am.")
        return ""

    try:
        command = recognizer.recognize_google(audio, language="en-IN")
        print(f"You: {command}")
        return command.lower().strip()

    except sr.UnknownValueError:
        print("I couldn't understand that.")
        return ""

    except sr.RequestError:
        speak("I'm having trouble connecting to speech recognition, ma'am.")
        return ""


# -------------------- COMMANDS --------------------

WEBSITES = {
    "youtube": ("https://www.youtube.com", "YouTube"),
    "google": ("https://www.google.com", "Google"),
    "github": ("https://github.com", "GitHub"),
    "gmail": ("https://mail.google.com", "Gmail"),
    "linkedin": ("https://www.linkedin.com", "LinkedIn"),
    "stackoverflow": ("https://stackoverflow.com", "Stack Overflow"),
    "stack overflow": ("https://stackoverflow.com", "Stack Overflow"),
}


def web_search(query):
    if query:
        speak(f"Searching the web for {query}, ma'am.")
        url = (
            "https://www.google.com/search?q="
            + urllib.parse.quote_plus(query)
        )
        open_url(url)
    else:
        speak("What would you like me to search for, ma'am?")


def _normalize(text):
    """Lowercase and remove spaces/symbols so 'playlist 1' == 'playlist1'."""
    return re.sub(r"[^a-z0-9]", "", text.lower())


def find_in_library(query):
    """Return (song_name, url) from the music library, or None."""
    q = _normalize(query)
    print(f"[music] Looking for: '{query}'  (normalized: '{q}')")

    if not q or not music:
        return None

    normalized = {_normalize(name): name for name in music}

    # 1. Exact match
    if q in normalized:
        name = normalized[q]
        print(f"[music] Exact match: {name}")
        return name, music[name]

    # 2. One contains the other ("tu janna song")
    for key, name in normalized.items():
        if len(key) >= 3 and (key in q or q in key):
            print(f"[music] Partial match: {name}")
            return name, music[name]

    # 3. Close match, to forgive speech-recognition mistakes
    close = difflib.get_close_matches(q, list(normalized), n=1, cutoff=0.5)
    if close:
        name = normalized[close[0]]
        print(f"[music] Close match: {name}")
        return name, music[name]

    print("[music] No match found.")
    return None


def play_music(query):
    """Play a song from the library, or search YouTube if it isn't there."""
    # Remove filler words: "play despacito on youtube" -> "despacito"
    query = re.sub(r"\bon youtube\b", "", query)
    query = re.sub(r"^(the )?(song|music|video)\s+", "", query)
    query = re.sub(r"\s+(song|music)$", "", query).strip()

    if not query:
        speak("Which song would you like me to play, ma'am?")
        query = listen()
        if not query:
            speak("I didn't hear a song name, ma'am.")
            return

    # 1. Your own music library
    match = find_in_library(query)
    if match:
        name, url = match
        speak(f"Playing {name}, ma'am.")
        open_url(url)
        return

    # 2. Not in the library: use YouTube
    speak(
        f"I couldn't find {query} in your library, ma'am. "
        "Searching YouTube for it."
    )
    print("Songs in library: " + ", ".join(music))

    if HAS_PYWHATKIT:
        try:
            pywhatkit.playonyt(query)  # opens in Chrome via webbrowser registration
            return
        except Exception as error:
            print(f"pywhatkit error: {error}")

    url = (
        "https://www.youtube.com/results?search_query="
        + urllib.parse.quote_plus(query)
    )
    open_url(url)


def handle_command(command):
    """Execute commands and speak the responses.
    Returns False when Jarvis should shut down, otherwise True."""

    command = command.lower().strip()
    words = re.findall(r"[a-z']+", command)  # whole words only

    # Play music (checked first so a song called "Stop" won't exit Jarvis)
    if command.startswith("play "):
        play_music(command[len("play "):].strip())
        return True

    if command in ("play", "play music", "play a song", "play song"):
        play_music("")
        return True

    # Exit Jarvis
    # (The wake word "jarvis" is removed before this runs, so
    # "stop jarvis" arrives here as just "stop".)
    if (
        command in ("stop", "exit", "quit", "goodbye")
        or any(w in words for w in ("exit", "quit", "goodbye"))
        or "shut down" in command
    ):
        speak("Goodbye, ma'am. Call me whenever you need me.")
        return False

    # Open websites
    for site, (url, name) in WEBSITES.items():
        if f"open {site}" in command or command == site:
            speak(f"Of course, ma'am. Opening {name} now.")
            open_url(url)
            return True

    # Search the web
    if command.startswith("search for "):
        web_search(command[len("search for "):].strip())
        return True

    if command.startswith("search "):
        web_search(command[len("search "):].strip())
        return True

    # Conversation
    if "how are you" in command:
        speak(
            "I'm doing well, ma'am. Thank you for asking. "
            "How can I help you?"
        )
        return True

    if "your name" in command or "who are you" in command:
        speak("I'm Jarvis, your personal voice assistant.")
        return True

    if "thank" in words or "thanks" in words or "thank you" in command:
        speak("You're welcome, ma'am.")
        return True

    if (
        any(w in words for w in ("hello", "hi", "hey"))
        or "good morning" in command
        or "good afternoon" in command
        or "good evening" in command
    ):
        speak("Hello, ma'am. How can I help you?")
        return True

    # Time
    if "time" in words:
        current_time = datetime.datetime.now().strftime("%I:%M %p")
        speak(f"It's {current_time}, ma'am.")
        return True

    # Date
    if "date" in words or "today" in words or "today's" in words:
        current_date = datetime.datetime.now().strftime("%A, %d %B %Y")
        speak(f"Today is {current_date}, ma'am.")
        return True

    # Unknown commands
    speak(
        "I'm not sure how to do that yet, ma'am. "
        "You can ask me to play a song, open YouTube, Google, GitHub, "
        "Gmail, or LinkedIn, search the web, tell you the time or date, "
        "or have a simple conversation."
    )
    return True


# -------------------- SINGLE INSTANCE GUARD --------------------

_instance_lock = None


def ensure_single_instance():
    """Stop Jarvis from running twice at the same time.
    Two copies would both hear you and both act on every command."""
    global _instance_lock
    _instance_lock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        _instance_lock.bind(("127.0.0.1", 47653))
    except OSError:
        print("Jarvis is already running in another window.")
        print("Close it first (or run: taskkill /F /IM python.exe).")
        sys.exit(1)


# -------------------- MAIN LOOP --------------------

def main():
    ensure_single_instance()

    print("=" * 48)
    print("          JARVIS VOICE ASSISTANT")
    print("=" * 48)
    print(f"Process ID: {os.getpid()}")
    print("Say 'Jarvis' to wake me up.")
    print("Say 'Jarvis, play <song name>' to play music.")
    print("Say 'Jarvis, stop' to exit.")

    # Prepare the microphone
    try:
        with sr.Microphone() as source:
            print("\nCalibrating microphone. Please stay quiet...")
            recognizer.adjust_for_ambient_noise(source, duration=1)

    except OSError as error:
        print(f"Microphone error: {error}")
        speak("I cannot access your microphone, ma'am.")
        return

    speak("Hello, ma'am. Jarvis is online and ready.")

    while True:
        command = listen()

        if not command:
            continue

        # Wake word
        if "jarvis" not in command:
            print("Say Jarvis to get my attention.")
            continue

        # Supports both "Jarvis" then a command,
        # and "Jarvis, open YouTube" in one go.
        user_command = command.replace("jarvis", "", 1).strip(" ,.!?")

        # Only ask "what can I do" if no command came with the wake word.
        if not user_command:
            speak("Yes, ma'am. What can I do for you?")
            user_command = listen()

        if not user_command:
            speak("I didn't hear a command, ma'am.")
            continue

        if not handle_command(user_command):
            break

        # Short pause so Jarvis doesn't hear his own voice or the music
        time.sleep(0.5)


if __name__ == "__main__":
    main()