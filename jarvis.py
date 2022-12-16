"""Author: K. Koyal
Date: 12/15/2022
E-mail : kkoyal19599@gmail.com
Virtual Assistant named: Jarvis"""

#modules required
import pyttsx3
import json
import requests
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import random
import psutil
from win32com.client import Dispatch
from bs4 import BeautifulSoup


#To get the voice
engine=pyttsx3.init('sapi5') 
voices=engine.getProperty('voices')
print(voices[0].id)
engine.setProperty('voice',voices[0].id)


#Commands to say while in process
lst_ans=["Command Accepted","Processing ma'am","okay ma'am","Your mish is my command"]
lst_app=["My pleasure ma'am","It's my job ma'am","It makes me happy to help","I'm always here to help","Anytime ma'am","I am here to serve"]


#To speak the assistant 
def speak(audio):
    engine.say(audio)
    engine.runAndWait()

#Greetings before taking any command
def wishMe():
    hour=int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Moring!")

    elif hour>=12 and hour<18:
        speak("Good Afternon!")
        
    else:                                                                                                                                                                                         
        speak("Good Evening!") 

    speak("I am Jarvis ma'am, Please tell me how may I help you")

#Taking command from the user in a voice
def takeCommand():
    # It takes microphone input from the user and returns string output
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1 #Time to wait before i complete saying energy_threshold is another var to increse the voice of yours to give command
        audio=r.listen(source)
    try:
        print("Recognizing...")
        query=r.recognize_google(audio,language='en-in')
        print(f"user said : {query}")
    except Exception as e:
        print(e)
        speak("Say that again please...")
        return "None"   #This is not a python none
    return query

#Closing all the opened tabs only
def close_app1(app_name):
        running_apps=psutil.process_iter(['pid','name']) #returns names of running processes
        found=False
        for app in running_apps:
            sys_app=app.info.get('name').split('.')[0].lower()

            if sys_app in app_name.split() or app_name in sys_app:
                pid=app.info.get('pid') #returns PID of the given app if found running
                
                try: #deleting the app if asked app is running.(It raises error for some windows apps)
                    app_pid = psutil.Process(pid)
                    app_pid.terminate()
                    found=True
                except: pass
                
            else: pass
        if not found:
            print(app_name+" not found running")
        else:
            print(app_name+'('+sys_app+')'+' closed')


#Newspaper reads for you
def speak1(string,i):
    speak=Dispatch("SAPI.spVoice")
    if num<10:
        speak.Speak(num)
    else:
        speak.Speak("And the last one is...")
    speak.Speak(string)


#main function starts here
if __name__=="__main__":
    wishMe()
    while True:
        query = takeCommand().lower()
        
        speak(random.choice(lst_ans))    

    #Logic for executing tasks
        #opening Wikipedia
        if 'wikipedia' in query:
            speak("Searching Wikipedia...")
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to wikipedia")
            print(results)
            speak(results)

        #opening youtube
        elif 'open youtube' in query:
            webbrowser.open("youtube.com")

        #opening Google
        elif 'open google' in query:
            webbrowser.open("google.com")

        #playing music from directory
        elif 'play music' in query:
            music_dir = "C:\\Users\\Win10\\Music\\" #c directory
            songs = os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir,random.choice(songs))) #selecting random songs from c directory

        #current time
        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Ma'am , the time is {strTime}")

        #opening vs code
        elif 'open code' in query:
            codePath = "C:\\Users\\Win10\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"    #It's a target of vs code not location of vscode
            os.startfile(codePath)

        #closing Google if open
        elif 'close google' in query:
            close_app1('chrome')

        #closing youtube if open
        elif 'close youtube' in query:
            close_app1('youtube')

        #closing vs code if open
        elif 'close code' in query:
            close_app1('code')

        #close music file if open
        elif 'close music' in query:
            close_app1('music')

        #Checking  current weather 
        elif 'weather' in query:
            
            
            # enter city name
            city = input("Enter the city name: ")
            
            # create url
            url = "https://www.google.com/search?q="+"weather"+city
            
            # requests instance
            html = requests.get(url).content
            
            # getting raw data
            soup = BeautifulSoup(html, 'html.parser')
            temp = soup.find('div', attrs={'class': 'BNeawe iBp4i AP7Wnd'}).text
            
            # this contains time and sky description
            str = soup.find('div', attrs={'class': 'BNeawe tAd8D AP7Wnd'}).text
            
            # format the data
            data = str.split('\n')
            time = data[0]
            sky = data[1]
            # list having all div tags having particular class name
            listdiv = soup.findAll('div', attrs={'class': 'BNeawe s3v9rd AP7Wnd'})
            
            # particular list with required data
            strd = listdiv[5].text
            
            # formatting the string
            pos = strd.find('Wind')
            other_data = strd[pos:]
            # printing all the data
            print("Temperature is", temp)
            print("Time: ", time)
            print("Sky Description: ", sky)
            print(other_data)
            speak("Temperature is ")
            speak(temp)
            speak("Time is ")
            speak(time)
            speak(" and Sky Description is ")
            speak(sky)
            speak(other_data)
            
        
        #Getting today's news
        elif "today's news" in query:
            speak=Dispatch("SAPI.spVoice")
            speak.Speak("Today's Hot Topics are")
            parameters = {
            'pageSize': 10,  # maximum is 100

        }
            url = ('YOUR API KEY')
            response = requests.get(url,params=parameters)
            p=response.json()
            num=1
            for i in p['articles']:
                string=i['title']
                print(string)
                print(i['url'],end="\n\n")
                speak1(string,num)
                num+=1
            speak.Speak("Thankyou for being with me")

        #opening timetable in this same directory
        elif "time table" in query:
            hour=int(datetime.datetime.now().hour)
            x = datetime.datetime.now().weekday()
            
            
            def SpeakingTimetable(fname,x):
                lines = open(fname).read().splitlines()
                print(lines[x])
                speak(lines[x])


            if hour>=14:
                speak("Tommorrow's Timetable")
                x = x+1
                SpeakingTimetable('table.txt',x)

            else:
                speak("Today's Timetable")
                SpeakingTimetable('table.txt',x)
                

        #response after thank you
        elif 'thank you' in query:
            speak(random.choice(lst_app))

        #Playing the game with computer snake, water and gun
        elif 'play game' in query:
            print("\t\t\t\t\tWelcome to 'snake Water Gun' Game\n")
            player_name=input("Enter your name: ")
            print("Start the Game.......\n")
            print(f"\t\t\t\t\t{player_name} VS Computer\n")
            lst=['Snake','Water','Gun']
            n=10
            i=1
            you,comp=0,0
            while n!=0:
    
                print(f"\nROUND {i}")
                select=input("\nEnter your choice: ")
                Select=select.capitalize()
                choice=random.choice(lst)
                if choice==Select:
                    you+=0
                    comp+=0
                    print("\n\t\t\t\t\tDraw")
                elif((choice=="Snake" and Select=="Gun") or (choice=="Gun" and Select=="Water") or (choice=="Water" and Select=="Snake")):
                    you+=1
                    comp+=0
                    print(f"\n\t\t\t\t\t{player_name} Wins")
                else:
                    you+=0
                    comp+=1
                    print("\n\t\t\t\t\tComputer Wins")
                print(f"YOU: {you}\t\t\t\t\t\t\t\t\t\t\tCOMPUTER: {comp}")
                n-=1
                i+=1
            print("\n\n\t\t\t\t\tGame Over")
            print("\t\t\t\t\tResults")
            print(f"YOU: {you}\t\t\t\t\t\t\t\t\t\t\tCOMPUTER: {comp}")
            if you>comp:
                print(f"\n\t\t\t\t\t{player_name} Winner")
            elif comp>you:
                print("\n\t\t\t\t\tComputer Winner")
            else:
                print("\n\t\t\t\t\tMatch Draws")
            print("\n\t\t\t\t\tThanks for Playing!!!")
        
        
         #closing the current open code (closing virtual assistant : jarvis)
        elif 'shutdown' in query:
            speak("Signing off")
            exit() 