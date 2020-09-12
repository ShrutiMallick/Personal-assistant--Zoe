#importing all the required packages
import subprocess
import wolframalpha      #to calculate strings into formulas
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import playsound            # to play saved mp3 files
import os                   # to save/open files
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
from clint.textui import progress
#from ecapture import ecapture as ec
from gtts import gTTS                 # google text to speech
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
from selenium import webdriver  #to control browsing operations

print('Loading your AI personal assistant-Sara')

#setting up our engine to pyttsx3

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour= int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good morning!")
    elif hour>=12 and hour<18:
        speak("Good Afternoon!")
    else:
        speak("Good evening!")

    assname=("Sara 1 point o")
    speak("I am your Assistant")
    speak(assname)





def usrname():
    speak("What should i call you")
    uname = takeCommand()
    speak("Welcome")
    speak(uname)
    columns =shutil.get_terminal_size().columns

    print("Welcome ",uname. center(columns))

    speak("How can i help you?")


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source,phrase_time_limit=5)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language='en-US')
            print(f"User said:{query}\n")
        except Exception as e:
            print(e)
            speak("Pardon me, please say it again!")
            return "None"
        return query

def sendEmail(to,content):
    server=smtplib.SMTP('smtp.gmail.com',587)
    server.ehlo()
    server.starttls()

#Enable low security in gmail
    server.login('your email id','your email password')
    server.sendmail('your email id',to ,content)
    server.close()

#Main function

if __name__=='__main__':
    clear=lambda :os.system('cls')      #This function will clear
                                        # any command before execution of this python file
    clear()
    wishMe()
    usrname()
while True:
    speak("How can i help you")
    query=takeCommand().lower()
    if query==0:
        continue
    if "good bye" in query or "ok bye" in query or"stop" in query:
        speak("Good bye,have a nice day.")
        print('Your assistant Sara is shutting down,Good Bye')
        break


    if 'wikipedia' in query:
        speak('Searching Wikipedia...')
        query=query.replace("wikipedia", " ")
        results=wikipedia.summary(query,sentences=3)
        speak("According to Wikipedia")
        print(results)
        speak(results)

    elif 'open youtube' in query:
        speak("Here you go to Youtube\n")
        webbrowser.open_new_tab("https://www.youtube.com")
        time.sleep(5)

    elif 'open google' in query:
        speak("Here you go to Google\n")
        webbrowser.open_new_tab("https://www.google.com")
        time.sleep(5)

    elif 'open gmail' in query:
        speak("Here you go to Gmail\n")
        webbrowser.open_new_tab("https://www.gmail.com")
        time.sleep(5)

    elif 'play music' in query or 'play song' in query:
        speak("Here you go with music")
        music_dir="C:\Users\win10\Music"
        songs=os.listdir(music_dir)
        print(songs)
        random=os.startfile(os.path.join(music_dir,songs[1]))

    





