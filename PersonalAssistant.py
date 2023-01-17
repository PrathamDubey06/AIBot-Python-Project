import os
import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import calendar
import smtplib
import requests
import pyautogui
import keyboard as kk
import time
import json
from win32com.client import Dispatch
import random as rand
import subprocess as sp
import pywhatkit as kit

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def greet():
    hour = datetime.datetime.now().hour
    if hour >= 0 and hour < 12:
        print(f"Good Morning")
        speak(f"Good Morning")
    elif hour >= 12 and hour < 17:
        print(f"Good Afternoon")
        speak(f"Good Afternoon")
    else:
        print(f"Good Evening")
        speak(f"Good Evening")


def greet_with_name(nm):
    hour = datetime.datetime.now().hour
    if hour >= 0 and hour < 12:
        print(f"Good Morning {nm}")
        speak(f"Good Morning {nm}")

    elif hour >= 12 and hour < 17:
        print(f"Good Afternoon {nm}")
        speak(f"Good Afternoon {nm}")
    else:
        print(f"Good Evening {nm}")
        speak(f"Good Evening {nm}")


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening... ")
        r.pause_threshold = 0.5
        audio = r.listen(source)
        try:
            print("Recognizing... ")
            query = r.recognize_google(audio, language='en-in')
            print(f"User Said: {query}")
        except Exception as e:
            print("Sorry, I couldn't get you. Could you say that again please?")
            speak("Sorry, I couldn't get you. Could you say that again please?")
            return "None"
    return query


def checkEmailName(name):
    dict = {
        "Name": "EmailId"
    }
    if name in dict:
        return (dict[name])
    return "Not Found"


def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('YourEmail', 'Password')
    server.sendmail('YourEmail', to, content)
    server.close()

def open_camera():
    sp.run('start microsoft.windows.camera:', shell=True)

def check_whatsapp(con_name):
    whatsapp_num={
        "Name":"Number"
    }
    if con_name in whatsapp_num:
        return (whatsapp_num[con_name])
    return "Not Found"

def send_whatsapp_message(contact_name, message):
    kit.sendwhatmsg_instantly(f"+91{contact_name}", message)
    pyautogui.click(1050, 950)
    time.sleep(5)
    kk.press_and_release('enter')


def send_whatsapp_message_with_Number(contactnum, message):
    kit.sendwhatmsg_instantly(f"+91{contactnum}", message)
    pyautogui.click(1050, 950)
    time.sleep(5)

    kk.press_and_release('enter')


#-------------------Chit_Chat Dictionary-------------------------
def chit_chat(what):
    hellos=[
        f'Hello {name}',
        f"Hola {name}",
        f"Bonjour {name}",
        f"nǐn hǎo"]

    hrus=[
        "Thanks for asking, Doing Well",
        "Things are good, I just had my second covid 19 shot",
        "It has been a rough week, You know lot of stuff to do..",
        "I have had a whirl wind of a week, but i am hanging in there",
        "Much better now, That you are with me",
        "The best i can be, assuming you are at your best too",
        "Different day, Same existence"
    ]

    jokes=[
        f"What's red and move up and down?  A tomato in an elevator",
        f"What do you call beers with no ears  B.. ha.. ha.. ha..",
        f"Why do french people eat snails.. They don't like fast food.. ha.. ha.. ha..",
        f"Want to hear a construction joke? , Oh never mind I am still working on that one..",
        f"How do you throw a space party?  You planet",
        f"Why am I slow?  because I am carrying a big burden like you.",
        f"How do the ocean says hi?  It waves",
        f"Wait.. Wait.. I found a good one. What does the storm cloud wear under his rain coat?  A thunderwear",
        f"Why did the teddy bear say no to dessert time? because the was stuffed"
        f"What do you call a guy who is really loud?  Mike ha.. ha.. ha.."
    ]

    introduce = [
        f"My mother name is: {myName}. \n My father's name is {myName} \n My brother's name is {myName} \n and I belong to {myName} family",
        f"I really lacked the words to compliment myself today. So, Nevermind",
        f"And Please Don't underestimate me. That's my mother's job",
        f"I am a nobody, Nobody is perfect. Therefore,  Yes you guessed it right I am perfect.."
    ]
    if what=="hru":
        ch1= rand.choice(hrus)
        print(ch1)
        speak(ch1)
    if what=="introduce":
        ch2 = rand.choice(introduce)
        print(ch2)
        speak(ch2)
    if what=="hello":
        ch3 = rand.choice(hellos)
        print(ch3)
        speak(ch3)
    if what=="joke":
        ch4 = rand.choice(jokes)
        print(ch4)
        speak(ch4)

if __name__ == '__main__':
    myName = "Jarvis"
    greet()
    print(f"I am {myName}")
    speak(f"I am {myName}")
    print("Can I get your name?")
    speak("Can I get your name?")
    p=True
    while p==True:
        choice = takeCommand().lower()
        if choice == "yes":
            print("What is your name: ")
            speak("What is your name: ")
            name = takeCommand()
            if name == "None":
                name = takeCommand()
            greet_with_name(name)
            print(f"How may I help you {name}?")
            speak(f"How may I help you {name}?")
            p=False
        elif choice=="no":
            print("No problem, I respect your privacy")
            speak("No problem, I respect your privacy")
            print(f"How may I help you?")
            speak(f"How may I help you?")
            name = "user"
            p=False
        else:
            speak("I don't know this command, Please say yes or no!")
            p=True
    k=True
    while k==True:
        query = takeCommand().lower()

        # ----------Wikipedia-------------
        if 'wikipedia' in query:
            print("Searching Wikipedia...")
            speak("Searching Wikipedia...")
            query = query.replace('wikipedia', '')
            result = wikipedia.summary(query, sentences=2)
            print("According to wikipedia...")
            speak("According to wikipedia...")
            print(f"\t{result}")
            speak(result)
            k=True

        # ---------Youtube---------------

        elif 'open youtube' in query:
            webbrowser.open("youtube.com")
            k = True

        elif 'search' and 'youtube' in query:
            q = query.replace('search', "")
            q = q.replace('youtube', " ")
            if 'on' in q:
                q = q.replace('on', '')
            print(f"Opening {q} on youtube")
            speak(f"Opening {q} on youtube")
            webbrowser.open(f"https://www.youtube.com/results?search_query={q}")
            k = True

        # -------Google----------

        elif 'open google' in query:
            webbrowser.open("google.com")
            k = True

        elif 'search' and 'google' in query:
            q = query.replace('search', "")
            q = q.replace('google', " ")
            if 'on' in q:
                q = q.replace('on', '')
            print(f"Opening {q} on google")
            speak(f"Opening {q} on google")
            webbrowser.open(f"https://www.google.com/search?q={q}")
            k = True

        # ------Current Time-------


        elif 'the' and 'time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"The time is: {strTime}")
            print(f"The time is: {strTime}")
            k = True

        elif 'the' and 'date' in query:
            strDate = datetime.datetime.now().date()
            speak(f"The date is: {strDate}")
            print(f"The date is: {strDate}")
            k = True

        elif "today's day" in query:
            d = str(datetime.datetime.now().date())
            s = []
            s = d.split("-")
            c_date = f"{s[2]} {s[1]} {s[0]}"
            born = datetime.datetime.strptime(c_date, '%d %m %Y').weekday()
            print(f"Today is: {calendar.day_name[born]}")
            speak(f"Today is: {calendar.day_name[born]}")
            k = True

        elif 'current weather' in query:
            d=query.replace("tell me the current weather of","")
            city_name=d.replace(" ",'')
            APIKEY = 'Your API Key'
            base_url = "http://api.openweathermap.org/data/2.5/weather?"
            complete_url = base_url + "appid=" + APIKEY + "&q=" + city_name
            response = requests.get(complete_url)
            x = response.json()
            if x["cod"] != "404":
                y = x["main"]
                current_temperature = round(y["temp"]-273,2)
                current_pressure = y["pressure"]
                current_humidity = y["humidity"]
                z = x["weather"]
                weather_description = z[0]["description"]
                print(" Temperature (in kelvin unit): " +
                      str(current_temperature) + "degree celsius"
                      "\n atmospheric pressure (in hPa unit): " +
                      str(current_pressure) +
                      "\n humidity (in percentage): " +
                      str(current_humidity) +
                      "\n description = " +
                      str(weather_description))
                speak(" Temperature (in kelvin unit): " +
                      str(current_temperature) +
                      "\n atmospheric pressure (in hPa unit):" +
                      str(current_pressure) +
                      "\n humidity (in percentage): " +
                      str(current_humidity) +
                      "\n description: " +
                      str(weather_description))
                k=True
            else:
                print(" City Not Found ")
                speak("Sorry, I cannot find the given city.")
                k=True

        elif 'current temperature' in query:
            d=query.replace("tell me the current temperature of","")
            city_name=d.replace(" ",'')
            APIKEY = 'Your API key'
            res = requests.get(
                f"http://api.openweathermap.org/data/2.5/weather?q={city_name}&appid={APIKEY}&units=metric").json()
            weather = res["weather"][0]["main"]
            temperature = res["main"]["temp"]
            feels_like = res["main"]["feels_like"]
            print(f"Weather: {weather} \n"
                  + f"Temperature: {temperature} ℃ \n" +
                  f"Feels like: {feels_like} ℃")
            speak(f"Weather: {weather}\n "
                  + f"Temperature: {temperature} ℃ \n" +
                  f"Feels like: {feels_like} ℃")

        # -------Send Email---------------

        elif 'email' in query:
            name = query.replace("email", "")
            name = name.replace("to", "")
            name = name.replace("send", "")
            name = name.replace(" ", "")
            email_addr = checkEmailName(name)
            if email_addr == "Not Found":
                try:
                    print("I dont have their data in my system")
                    speak("I dont have their data in my system")
                    speak("Please provide me the email address:")
                    to = takeCommand().lower()
                    to = to.replace("at the rate", "@")
                    to = to.replace(" ", "")
                    print("What should I send?")
                    speak("What should I send?")
                    content = takeCommand()
                    sendEmail(to, email_addr)
                    print("Email has been sent!")
                    speak("Email has been sent!")
                    k = True
                except Exception as e:
                    print(e)
                    print(f"Sorry {name}, I am unable to send email!")
                    speak(f"Sorry {name}, I am unable to send email!")
                    k = True
            else:
                try:
                    print("What should I send?")
                    speak("What should I send?")
                    content = takeCommand()
                    to = email_addr
                    sendEmail(to, content)
                    print("Email has been sent!")
                    speak("Email has been sent!")
                    k = True
                except Exception as e:
                    print(e)
                    print(f"Sorry {name}, I am unable to send email!")
                    speak(f"Sorry {name}, I am unable to send email!")
                    k = True

        #-------Whatsapp Message------
        elif 'send' and 'whatsapp' in query:
            nm=query.replace("send","")
            nm=nm.replace("whatsapp","")
            nm=nm.replace("message","")
            nm=nm.replace("to","")
            nm=nm.replace("on","")
            nm=nm.replace(" ",'')
            whatsappNumber=check_whatsapp(nm)
            if whatsappNumber=="Not Found":
                print("This name is not registered on my system. Can you please tell me their number")
                speak("This name is not registered on my system. Can you please tell me their number")
                whatsappNumber=takeCommand()
                print("What do you want me to send?")
                speak("What do you want me to send?")
                message=takeCommand()
                send_whatsapp_message_with_Number(whatsappNumber,message)
                speak(f"I have sent the message to {whatsappNumber}!")
                k=True
            else:
                print("What do you want me to send?")
                speak("What do you want me to send?")
                message=takeCommand()
                send_whatsapp_message(whatsappNumber,message)
                speak(f"I have sent the message to {nm}")
                k=True

        #-------Today's News----------
        elif "today's news" in query:
            apikey="Your API Key"
            r = requests.get(f"https://newsapi.org/v2/top-headlines?country=in&apiKey={apikey}")
            res = r.json()
            print(res['status'])

            print(f"Hello {name}")
            speak(f"Hello {name}")
            print("Welcome to your favourite news channel")
            speak("Welcome to your favourite news channel")
            print("Today's top 5 news are")
            speak("Today's top 5 news are")

            news = res['articles']

            for j in range(1, 6):
                i = rand.randint(1, 20)
                if j <= 4:
                    if news[i]['description'] == None:
                        print(j)
                        speak(j)
                        print(f"Title: \n\t {news[i]['title']}")
                        speak(f"The title is : {news[i]['title']}")
                        print(f"Description: \n\t {news[i]['description']}")
                        speak(f"There is no description provided for the above news")
                    else:
                        print(j)
                        speak(j)
                        print(f"Title: \n\t {news[i]['title']}")
                        speak(f"The title is : {news[i]['title']}")
                        print(f"Description: \n\t {news[i]['description']}")
                        speak(f"The description is: {news[i]['description']}")
                else:
                    speak("And the last news for today is: ")
                    print(j)
                    speak(j)
                    print(f"Title: \n\t {news[i]['title']}")
                    speak(f"The title is : {news[i]['title']}")
                    print(f"Description: \n\t {news[i]['description']}")
                    speak(f"The description is: {news[i]['description']}")
            k=True

        #-----Open Camera-------------

        elif 'open camera' in query:
            speak("Opening camera..")
            open_camera()
            k=True

        # ------Open Applications----------
        elif 'open' in query:
            q = query.replace('open', '')
            q = q.replace(" ", "")
            dict = {
                "zoom": "C:\\Users\\prath\\AppData\\Roaming\\Zoom\\bin\\Zoom.exe",
                "vscode": "C:\\Users\\prath\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe",
                "pycharm": "C:\\Program Files\\JetBrains\\PyCharm Community Edition 2021.3.1\\bin\\pycharm64.exe",
                "eclipse": "C:\\EclipseEE\\eclipse\\eclipse.exe",
                "scilab": "C:\\Program Files\\scilab-6.1.1\\bin\\WScilex.exe",
                "discord": "C:\\Users\\prath\\AppData\\Local\\Discord\\Update.exe --processStart Discord.exe"
            }
            cnt=0
            for app in dict:
                if (app in q):
                    speak(f"Opening {q}")
                    os.startfile(dict[q])
                    k = True
                else:
                    cnt=cnt+1
            if cnt==len(dict):
                speak(f"Sorry {name}, I cannot open {q}.")
                k=True


        # ---------What Can I Do----------
        elif 'what' and 'you do' in query:
            speak(f"Hmmmmmmmmm...That's a great question {name}")
            speak("Let me show you what all I can do to ease your life.")
            I_Can_Do={
                "Wikipedia" : "I can increase your knowledge by searching anything you want on wikipedia",
                "Google" : "I can open google and search anything you want me to.. but don't ask me to go to incognito :P",
                "Youtube" : "I can play any videos on youtube",
                "News" : "I can be your news reporter and tell you daily top 5 news",
                "Today's Everything" : "I can tell you current time, day, date as well as current temperature and weather so that you won't miss taking umbrella with you xD",
                "Open Applications" : "I can open various applications that are in your system",
                "Email": " I can send email to anyone using your email id without even letting you touch your keyboard",
                "Whatsapp" : "I can chat with your friends on your behalf, You just need to say the message and Its my duty to deliver it",
                "Jokes" : "I can tell you the best jokes in this universe. ha..ha..ha.. I am already laughing"
            }
            for key,value in I_Can_Do.items():
                print(key," : ",value)
                speak(f"{key} : {value}")
            k=True
        # ----Some Chit Chat------------
        elif 'how are you' in query:
            hrus = [
                "Thanks for asking, Doing Well",
                "Things are good, I just had my second covid 19 shot",
                "It has been a rough week, You know lot of stuff to do..",
                "I have had a whirl wind of a week, but i am hanging in there",
                "Much better now, That you are with me",
                "The best i can be, assuming you are at your best too",
                "Different day, Same existence"
            ]
            obj = rand.choice(hrus)
            print(obj)
            speak(obj)
            obj=""
            k = True


        elif 'tell me a joke' in query:
            jokes = [
                f"What's red and move up and down?  A tomato in an elevator",
                f"What do you call beers with no ears  B.. ha.. ha.. ha..",
                f"Why do french people eat snails.. They don't like fast food.. ha.. ha.. ha..",
                f"Want to hear a construction joke? , Oh never mind I am still working on that one..",
                f"How do you throw a space party?  You planet",
                f"Why am I slow?  because I am carrying a big burden like you.",
                f"How do the ocean says hi?  It waves",
                f"Wait.. Wait.. I found a good one. What does the storm cloud wear under his rain coat?  A thunderwear",
                f"Why did the teddy bear say no to dessert time? because the was stuffed"
                f"What do you call a guy who is really loud?  Mike ha.. ha.. ha.."
            ]
            obj = rand.choice(jokes)
            print(obj)
            speak(obj)
            obj=""
            k = True

        elif  'hello' in query:
            # chit_chat('hello')
            hellos = [
                f'Hello ',
                f"Hola ",
                f"Bonjour",
                f"nǐn hǎo"]
            obj = rand.choice(hellos)
            print(obj)
            speak(obj)
            obj=""
            k = True

        elif f'introduce {myName}' in query:
            introduce = [
                f"My mother name is: {myName}. \n My father's name is {myName} \n My brother's name is {myName} \n and I belong to {myName} family",
                f"I really lacked the words to compliment myself today. So, Nevermind",
                f"And Please Don't underestimate me. That's my mother's job",
                f"I am a nobody, Nobody is perfect. Therefore,  Yes you guessed it right I am perfect.."
            ]
            obj = rand.choice(introduce)
            print(f"My name is Jarvis {obj}")
            speak(f"My name is Jarvis {obj}")
            obj=""
            k = True

        #-------To Quit Jarvis-----------
        elif 'quit' in query:
            print(f"Its time to say goodbye. Thank you for using me {name}.")
            speak(f"Its time to say goodbye. Thank you for using me {name}.")
            k = False
