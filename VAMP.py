import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import random
import smtplib
import pandas as pd
import requests
import json
import threading

engine = pyttsx3.init('sapi5')
voice = engine.getProperty('voices')[1]  # get voice from microsoft text to speech
engine.setProperty('voice', voice.id)  # set above voice by default [0] is selected
engine.setProperty('rate', 170)  # set speed of voice, by default 200


def back(s):
    execute(s)


def speak(sp):
    """Speak the string which have been passed as an argument"""
    try:
        engine.say(sp)
        engine.runAndWait()
    except RuntimeError:
        print("Unable to start")


def wishMe():
    """Wishes user good morning, afternoon, evening according to system time, and tells about itself"""
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 12:
        speak("Good Morning!")
    elif 12 <= hour < 18:
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")
    speak("I am VAMP. Please tell me how may i help you?")  # Virtual Assistant by Meet Pogul


def takeCommand():
    """Take Command From User and do speech recognition"""

    r = sr.Recognizer()  # set recognizer

    with sr.Microphone() as source:
        print("Listening...")
        r.adjust_for_ambient_noise(source)  # auto energy threshold according to noise outside
        r.pause_threshold = 1  # if it takes 1 sec gap it will not complete session of listening
        # r.energy_threshold = 250  # it will check user voice at that level
        audio = r.listen(source)
    try:
        print("Recognizing...")
        q1 = r.recognize_google(audio, language="en-in")  # recognize voice and convert into text
        print(f"User: {q1}")
    except Exception as e:
        print(e)
        print("Please say it again...")
        q1 = takeCommand()  # if error occur again take command
    return q1


def weather():
    """Gives weather Report Through Open Weather Map using json"""
    with open(r"D:\Meet\Python\firstprog\VAMP\OPM.txt", "r") as f:  # get api from file
        api = f.readline()
        api = api.replace("\n", "")
    url = requests.get(
        fr"http://api.openweathermap.org/data/2.5/weather?id=1255364&appid={api}")
    # get json from website
    content = json.loads(url.text)  # convert into dictionary
    # read weather
    w = content['weather'][0]["description"]
    speak(f"Today's weather is {w}")
    c = (content["main"]["temp"] - 273.15)
    speak(f"Temperature {c:.2f} degree celsius")
    hu = content["main"]["humidity"]
    speak(f"humidity {hu}%")
    clouds = content["clouds"]["all"]
    speak(f"clouds {clouds}%")
    pressure = content["main"]["pressure"]
    speak(f"Pressure {pressure} h.p.a")
    # wind = content["wind"]["speed"] * 2.237
    # speak(f"Wind speed {wind} miles per hour")


def water():
    """start file of water remainder in background"""
    os.system(r'pythonw D:\Meet\Python\firstprog\PracticePython6.py')


def mail():
    """Process on email take name match email not found then ask to write or cancel then continue, found or typed
    then ask for subject and content """

    def send_email(email, message):
        """Send email from folowing gmail Id"""
        GMAIL_ID = 'mspsoftwarex@gmail.com'
        with open(r"C:\Users\meets\OneDrive\Documents\infomail.txt", "r") as f:  # get password from file
            a = f.readline()
            a = a.replace("\n", "")
        GMAIL_PASWD = a
        s = smtplib.SMTP('smtp.gmail.com', 587)  # establish protocol
        s.ehlo()
        s.starttls()
        s.login(GMAIL_ID, GMAIL_PASWD)  # login
        s.sendmail(GMAIL_ID, email, message)  # send email
        s.close()

    try:
        speak("Whom to send")
        name = takeCommand()  # take name
        if name.lower() == 'cancel':
            print()
            return None
        to = ""
        df = pd.read_excel(r"D:\Meet\Python\firstprog\VAMP\emails.xlsx")  # read excel file
        for index, item in df.iterrows():
            if name.lower() in str(item["Name"]).lower():  # check name in list
                to = str(item["Email"])  # retrieve email id
        if to == "":  # if email not found in list
            speak("Do you want to continue by writing email id")
            y = takeCommand()  # take command to continue
            if y != 'yes' and y != 'yeah' and y != 'haa':
                speak("email canceled")
                print()
                return None
            to = input("Enter your Email Id: ")
        print("Subject: ")  # take subject
        speak("On which Subject")
        sub = takeCommand()
        if sub.lower() == 'cancel':  # to cancel whenever want
            print()
            return None
        print("Message:")
        speak("What you want to send")  # message
        msg = takeCommand()
        if sub.lower() == 'cancel':
            print()
            return None
        msg = f"Subject: {sub}\n\n {msg}\n\n\n\nMeet Pogul"
        send_email(to, msg)
        speak("Email sent")
    except Exception as e:
        print(e)
        speak("Email not sent, due to some technical issue")


def search(q2):
    """Search in respective website"""
    chrome = r"C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s"
    q2 = q2.replace("search ", "")
    q2 = q2.replace("Search ", "")
    if 'on google' in q2 or 'in google' in q2:  # search in google
        q2 = q2.replace("on google", "")  # remove unnecessary thing from query
        q2 = q2.replace("in google", "")
        webbrowser.get(chrome).open_new_tab("google.com/?#q=" + q2)
    elif 'on youtube' in q2 or 'in youtube' in q2:  # search in youtube
        q2 = q2.replace("on youtube", "")
        q2 = q2.replace("in youtube", "")
        webbrowser.get(chrome).open_new_tab("youtube.com/results?search_query=" + q2)
    elif 'on facebook' in q2 or 'in facebook' in q2:  # search in facebook
        q2 = q2.replace("on facebook", "")
        q2 = q2.replace("in facebook", "")
        webbrowser.get(chrome).open_new_tab("facebook.com/search/top?q=" + q2)
    elif 'on wikipedia' in q2 or 'in wikipedia' in q2:  # search in wikipedia
        q2 = q2.replace("on wikipedia", "")
        q2 = q2.replace("in wikipedia", "")
        webbrowser.get(chrome).open_new_tab(
            f"https://en.wikipedia.org/w/index.php?cirrusUserTesting=control&sort=relevance"
            f"&search={q2}&title=Special:Search&profile=advanced&fulltext=1&advancedSearch"
            f"-current=%7B%7D&ns0=1")


def execute(q3):
    """Compare query to following if elif command and execute inside code"""
    chrome = r"C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s"
    check = 'kholo' in q3 or 'open' in q3
    commands = ['search ... in google', 'search ... in youtube', 'search ... in wikipedia', 'search ... in facebook',
                'wikipedia', 'weather', 'open youtube', 'open google meet', 'open google', 'open whatsapp',
                'open classroom', 'open gmail', 'open python documentation', 'open stackoverflow', 'play music',
                'tell time', 'open jupyter notebook', 'open chrome', 'open t o r', 'open gmail', 'water reminder',
                'who are you', 'how are you', 'exit', 'python code', 'initialize day', 'list all commands',
                'open semester', 'open timetable', 'open vishruti mam website']

    if 'search' in q3:
        search(q3)

    elif ('initial' in q3 or 'start' in q3) and 'today' in q3:
        threading.Thread(target=back, args=['open whatsapp']).start()
        threading.Thread(target=back, args=['open gmail']).start()
        execute('water reminder')
        execute('tell date')
        execute('weather')

    elif 'wikipedia' in q3:
        try:
            q3 = q3.replace("wikipedia", "")
            result = wikipedia.summary(q3, sentences=2)
            print(result)
            speak("According to Wikipedia")
            speak(result)
        except wikipedia.DisambiguationError:
            speak("Could not found any particular, try search on wikipedia command")

    elif 'weather' in q3:
        weather()

    elif 'youtube' in q3 and check:
        webbrowser.get(chrome).open_new_tab("youtube.com")

    elif 'google meet' in q3 and check:
        webbrowser.get(chrome).open_new_tab("meet.google.com")

    elif 'google' in q3 and check:
        webbrowser.get(chrome).open_new_tab("google.com")

    elif 'whatsapp' in q3 and check:
        webbrowser.get(chrome).open_new_tab("web.whatsapp.com")

    elif 'classroom' in q3 and check:
        threading.Thread(target=back, args=["open vishruti mam website"]).start()
        webbrowser.get(chrome).open_new_tab("classroom.google.com")

    elif 'gmail' in q3 and check:
        webbrowser.get(chrome).open_new_tab("mail.google.com")

    elif ('python documentation' in q3 or 'python docs' in q3) and check:
        webbrowser.get(chrome).open_new_tab("docs.python.org/3/library/index.html")

    elif ('stack overflow' in q3 or 'stackoverflow' in q3) and check:
        webbrowser.get(chrome).open_new_tab("stackoverflow.com")

    elif 'play music' in q3 or 'music chalao' in q3:
        music_dir = r'E:\Music'
        songs = os.listdir(music_dir)
        os.startfile(os.path.join(music_dir, random.choice(songs)))

    elif ('time' in q3) and ('batavo' in q3 or 'tell' in q3 or 'till' in q3):
        strTime = datetime.datetime.now().strftime("%H:%M:%S")
        print(f"Current Time is, {strTime}")
        speak(f"Current Time is, {strTime}")

    elif ('date' in q3) and ('batavo' in q3 or 'tell' in q3 or 'till' in q3):
        strTime = datetime.datetime.now().strftime("%d-%m-%Y")
        print(f"Today's date is, {strTime}")
        speak(f"Today's date is, {strTime}")

    elif ('jupyter notebook' in q3 or 'jupiter notebook' in q3) and check:
        curr = os.getcwd()
        os.chdir(r'D:\Meet\Python')
        os.startfile('jupyter-notebook.exe')
        os.chdir(curr)

    elif 'chrome' in q3 and check:
        os.startfile(r'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe')

    elif ('t o r' in q3 or 't o r' in q3) and check:
        os.startfile(r'D:\Software\Tor Browser\Browser\firefox.exe')

    elif ('email' in q3 or 'gmail' in q3 or 'message' in q3) and ('send' in q3 or 'bhejo' in q3):
        mail()

    elif 'water reminder' in q3:
        t1 = threading.Thread(target=water)
        t1.start()
    ######################################################
    elif 'semester' in q3 and check:
        os.startfile(r'D:\Meet\College\Sem5')

    elif ('timetable' in q3 or 'time table' in q3) and check:
        os.startfile(r'D:\Meet\College\Sem5\BE-III_5 Sem_2W.jpg')

    elif ('vishruti' in q3) and check:
        webbrowser.get(chrome).open_new_tab("vishrutidesai.gnomio.com")
    ###################################################################
    elif q3 == "hello" or q3 == "hi" or "who are you" in q3 or q3 == "hai" or "kaun ho" in q3:
        speak("Hello, I am Vamp, i am your virtual assistant developed by meet pogul")

    elif "how are you" in q3 or "kaisa hai" in q3:
        speak("I am Fine, Thank you for asking")

    elif ('quit' in q3 or 'exit' in q3 or "band" in q3) and ('vamp' in q3 or 'wamp' in q3):  # quit
        speak("Thank you, Start me when you want to do any other stuff")
        exit()

    elif 'python code' == q3 or 'code python' == q3:  # necessary for python open
        print("hello")
        os.startfile(r'E:\Python CodeWithHarry')
        os.startfile(r'D:\Meet\Python')
        execute('open jupyter notebook')
        os.startfile(r'C:\Users\meets\AppData\Local\JetBrains\PyCharm 2020.1\bin\pycharm64.exe')

    elif 'list' in q3 and 'commands' in q3:  # list all command
        for i in commands:
            print(f"{commands.index(i)}. {i}")
            speak(f"{commands.index(i)}. {i}")

    print()


if __name__ == '__main__':  # main function
    wishMe()  # wish me whenever starts
    query = None
    while query != 'quit':
        query = takeCommand().lower()  # take command till quit
        execute(query)  # execute query

# more to do
# schedule task
# take report how much time user is lock
# learn above modules
# for text message to phone number
# new thread for every execute(), but if any execution contradict with another
