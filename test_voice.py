
import pyttsx3

engine = pyttsx3.init("sapi5")
engine.setProperty("rate", 170)
engine.setProperty("volume", 1.0)

print("Testing Jarvis voice...")
engine.say("Yes, ma'am. What can I do for you?")
engine.runAndWait()

print("Voice test finished.")
