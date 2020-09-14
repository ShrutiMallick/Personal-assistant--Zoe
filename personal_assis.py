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
        
        elif ' the time' in query:
        strTime=datetime.datetime.now().strftime("% H:% M:% S")
        speak("The time is {strTime}")

    elif 'send a mail' in query:
        try:
            speak("What should I say?")
            content=takeCommand()
            speak("Whom should i send")
            to=input()
            sendEmail(to,content)
            speak("Email has been sent!")
        except Exception as e:
            print(e)
            speak("I am not able to send this email")

    elif 'weather' in query:
        api_key="8ef61edcf1c576d65d836254e11ea420"
        base_url="https://api.openweathermap.org/data/2.5/weather?"
        speak("Whats the city name")
        city_name=takeCommand()
        complete_url=base_url+"appid="+api_key+"&q="+city_name
        response=requests.get(complete_url)
        x=response.json()
        if x["cod"]!="404":
            y=x["main"]
            current_temperature=y["temp"]
            current_humidity=y["humidity"]
            z=x["weather"]
            weather_description=z[0]["description"]
            speak("Temperature in kelvin unit is"+str(current_temperature)+"\n humidity in percentage is"+
                  str(current_humidity)+"\n description "+str(weather_description))
            print("Temperature in kelvin unit is"+str(current_temperature)+"\n humidity in percentage is"+
                  str(current_humidity)+"\n description "+str(weather_description))
        else:
            speak(" City Not Found ")

    elif "how are you" in query:
        speak("I am fine.Thank you")
        speak("How are you?")

    elif 'fine' in query or 'good' in query:
        speak("It's good to know that your fine")

    elif 'change name' in query:
        speak("What would you like to call me?")
        assname=takeCommand()
        speak("Thanks for namimg me")

    elif "What's your name" in query or "What is your name" in query:
        speak("My friends call me")
        speak(assname)
        print("My friends call me",assname)

    elif 'exit' in query:
        speak("Thanks for giving me your time")
        exit()

    elif "Who made you" in query:
        speak("I have been created by Shruti")

    elif "Who are you" in query or "What can you do" in query:
        speak('I am Zoe version 1 point o your personal virtual assistant.I am programmed to perform minor tasks like'
              'opening youtube,google chrome,gmail ,predict time,take a photo,search,predict weather in different '
              'cities,get top headline of news and i can also solve computational,geographical and mathematical '
              'questions also i can crack good jokes do you want to listen some?')

    elif "joke" in query or "yes i do" in query:
        speak(pyjokes.get_joke())
        print(pyjokes.get_joke())

    elif "calculate" in query:
        app_id="33V5YL-HYYU4PK5KL"
        client=wolframalpha.Client('33V5YL-HYYU4PK5KL')
        indx=query.lower().split().index('calculate')
        query=query.split()[indx+1:]
        res=client.query(' '.join(query))
        answer=next(res.results).text
        print("The answer is "+answer)
        speak("The answer is "+answer)

    elif 'search' in query or 'play':
        query=query.replace("search"," ")
        query=query.replace("play"," ")
        webbrowser.open(query)

    elif "who am i" in query:
        speak("If you talk then definately you are human")

    elif "Why you came to world" in query:
        speak("Thanks to Shruti further it's a secret!")

    elif "is love" in query:
        speak("It is the 7th sense that destroys all other senses")

    elif "reason for you" in query or "why were you made":
        speak("As shruti had no work to do in lockdown she made me as her personal assisstant"
              " and for some fun and also to learn new things")

    elif 'news' in query:
        news=webbrowser.open_new_tab("https://timesofindia.indiatimes.com/home/headlines")
        speak('Here are some headlines from the Times of India,Happy reading!')
        time.sleep(6)

    elif 'lock window' in query:
        speak("Locking the device")
        ctypes.windll.user32.LockWorkStation()

    elif 'shutdown system' in query:
        speak("Hold on a Sec!Your system is on its way to shutdown")
        subprocess.call('shutdown/p/f')

    elif 'empty recycle bin' in query:
        winshell.recycle_bin().empty(confirm=False,show_progress=False,sound=True)
        speak("Recycle Bin recycled!")

    elif "don't listen" in query or "stop listening" in query:
        speak("for how much time you want to stop Zoe from listening commands")
        a=int(takeCommand())
        time.sleep(a)

    elif "where is" in query:
        query=query.replace("Where is"," ")
        location=query
        speak("User asked to Locate")
        speak(location)
        webbrowser.open("https://www.google.nl//map/place/"+location+" ")

    #elif "camera" in query or "take a photo" in query or "take a pic":
       # ec.capture(0,"ZOe Camera","imag.jpg")

    elif "restart" in query:
        subprocess.call("shutdown","/r")

    elif "hibernate" in query or "sleep" in query:
        speak("Hibernating")
        subprocess.call("shutdown","/h")






    





