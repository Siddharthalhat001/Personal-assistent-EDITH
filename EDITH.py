import pyttsx
import speech_recognition as sr
import datetime
import os
import random
from requests import get
import wikipedia
import webbrowser
import pywhatkit as msg
import smtplib
import sys
import pyautogui
import time
import requests
from win32com import client



engine = pyttsx.init('sapi5')
voices = engine.getProperty('voices')
#print(voices[1].id)
engine.setProperty('voice', voices[len(voices)-1].id)

#text to speech
def speak(audio):
    engine.say(audio)
    engine.runAndWait()

    
def sendEmail(to,content):
    server = smtplib.SMTP('smtp.gmail.com',587)
    server.ehlo()
    server.starttls()
    server.login('your email id', 'password')
    server.sendmail("your email id", to ,content)
    server.close()

    
def news():
    main_url = "http://newsapi.org/v2/everything?q=tesla&from=2021-02-03&sortBy=publishedAt&apiKey=c1a67e3f9b774be79798104128cb777e"
    main_page = requests.get(main_url).json()
    #print main page
    articles = main_page['articles']
    #print (articles)
    head = []
    day = ['first', 'second', 'third','fourth', 'fifth', 'sixth']
    for ar in articles:
        head.append(ar['title'])
    for i in range(len(day)):
        speak(f"today's {day[i]} news  is: {head[i]} ")


#to convert voice into text
def takeCommand():
    r =sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source,timeout=1,phrase_time_limit=5)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio , language='en-in')
        print(f"user said : {query} ")
    except Exception as e:
        speak("Say that again please....")
        return  'none'
    return query

  
#to wish gm etc
def Wish():
    hour = int(datetime.datetime.now().hour)
    if hour >=0 and hour <=12:
        speak("Good Morning Siddharth")
    elif hour >12 and hour <=18:
        speak("Good Afternoon Siddharth")
    else:
        speak("Good Evening Siddharth")
    speak("I am Edith. Sir Please tell me how may i help you")



if __name__ == '__main__':

    # speak("Hi I am Siddhu. Sir How may I help you")
    #takeCommand()
    Wish()
    while True:
        if 1:
            query = takeCommand().lower()
            if 'open notepad' in query:
                npath = 'C:\\Windows\\System32\\Notepad.exe'
                os.startfile(npath)

                
            elif 'open cmd' in query:
                os.system("start cmd")

                
            elif 'open cmd' in query:
                os.system("start cmd")

            # elif "open camera" in query:
            #     cap

            
            elif "song" in query:
                music_dir = "D:\\Romantic songs"
                songs = os.listdir(music_dir)
                rd = random.choice(songs)
                os.startfile(os.path.join(music_dir, rd))
                speak('Okay, here is your music! Enjoy!')

                
            elif "ip address" in query:
                ip = get("https://api.ipify.org").text
                speak(f"Your IP address is {ip} ")

                
                
            elif "wikipedia" in query:
                speak("Searching wikipedia...")
                query = query.replace("wikipedia", "")
                results = wikipedia.summary(query,sentences=2)
                speak("according to wikipedia ")
                speak(results)
                print("according to wikipedia ")
                print(results)

            # elif "youtube" in query:
            #     webbrowser.open("www.youtube.com")

            
            elif "instagram" in query:
                webbrowser.open("www.instagram.com")

                
                
            elif "google" in query:
                speak("Siddharth .Sir, what should I search on Google")
                search = takeCommand().lower()
                webbrowser.open(f" {search} ")

                
                
            # elif "send message" in query:
            #     #speak("Sir, what message should i send")
            #     msg.sendwhatmsg("+919730829060", "Friday is running",2,36)

            
            
            elif "youtube" in query:
                msg.playonyt("moh moh ke dhage")
                # speak("Sir, which song  should i play on youtube")
                # search = takeCommand().lower()
                # webbrowser.open(f" {search} ")

                
                
            # elif "mail to sidd" in query:
            #     try:
            #         speak("what should i say ?")
            #         content = takeCommand().lower()
            #         to = "siddzzalhatgaming@gmail.com"
            #         sendEmail(to,content)
            #         speak("Email has been sent to Siddharth")
            #     except Exception as e:
            #         speak("Sorry to disappoint you Sir, but i am unable to send the message")
            #         print("Sorry to disappoint you Sir, but i am unable to send the message")

            
            
            elif 'email' in query:
                speak('Who is the recipient? ')
                recipient = takeCommand()

                if 'me' in recipient:
                  
                    try:
                        speak('What should I say? ')
                        content = takeCommand()

                        server = smtplib.SMTP('smtp.gmail.com', 587)
                        server.ehlo()
                        server.starttls()
                        server.login("Your_Username", 'Your_Password')
                        server.sendmail('Your_Username', "Recipient_Username", content)
                        server.close()
                        speak('Email sent!')

                    except:
                        speak('Sorry Sir! I am unable to send your message at this moment!')


            elif 'open gmail' in query:
                speak('okay')
                webbrowser.open('www.gmail.com')

                
            elif "what\'s up" in query or 'how are you' in query or 'how  you doing' in query:
                stMsgs = ['Just doing my thing!', 'I am fine!', 'Nice!', 'I am nice and full of energy']
                speak(random.choice(stMsgs))

                

            elif "switch the window" in query:
                pyautogui.keyDown("alt")
                pyautogui.press("tab")
                time.sleep(1)
                pyautogui.keyUp("alt")

                
            elif "news" in query:
                speak("Please wait Sir, Fetching the latest news")
                news()



                
            elif 'nothing' in query or 'abort' in query or 'stop' in query:

                speak('okay')

                speak('Bye Sir, have a good day.')

                sys.exit()

                
            #speak("sir, do you have any other work")

            # else:
            #     query = query
            #     speak('Searching...')
            #     try:
            #         try:
            #             res = client.query(query)
            #             results = next(res.results).text
            #             speak('WOLFRAM-ALPHA says - ')
            #             speak('Got it.')
            #             speak(results)
            #
            #         except:
            #             results = wikipedia.summary(query, sentences=2)
            #             speak('Got it.')
            #             speak('WIKIPEDIA says - ')
            #             speak(results)
            #
            #     except:
            #         webbrowser.open('www.google.com')
            #
            #
















