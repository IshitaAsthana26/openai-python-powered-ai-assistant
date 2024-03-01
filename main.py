import win32com.client
import pyttsx3
import speech_recognition as sr
import os

engine = pyttsx3.init()


speaker = win32com.client.Dispatch("SAPI.Spvoice")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
    try:
         query = r.recognize_google(audio,language="hi-in")
         print(f"User said; {query}")
         return query
    except Exception as e:
         return "Some Error Occured. Sorry from Jarvis"

if __name__ == '__main__':
    print('PyCharm')
    engine.say("Hello I am Jarvis AI")
    engine.runAndWait()

    while True:

     print("Listening...")
     query = takeCommand()




