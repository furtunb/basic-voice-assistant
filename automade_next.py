import speech_recognition as sr
import pyttsx3
import pywhatkit
import pyjokes
import wikipedia
import pyautogui
import time
# from win32com.client import Dispatch # I didn't use it because I liked pyttsx3 more than this

""" I pyttsx3 is better
def speak():
    speak = Dispatch("SAPI.SpVoice")
    speak.speak("Hello there, I am computer")
"""

screenWidth, screenHeight = pyautogui.size()
currentMouseX, currentMouseY = pyautogui.position()

listener = sr.Recognizer()
engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)


def speak(text):
    engine.say(text)
    engine.runAndWait()


def listen():
    name = "alexa"
    speak("Hi there, how can I help you")
    try:
        with sr.Microphone() as mic:
            print("you are using my microphone")
            voice = listener.listen(mic)
            request = listener.recognize_google(voice)
            print(request + " I listened you:)")
            request = request.lower()
            if name in request:
                request = request.replace(name, '')
                print(request)
            else:
                speak("You didn't call my name")
    except:
        request = 0
    print(request)
    return request


def listen_to_verify(lang="en-US"):
    try:
        with sr.Microphone() as mic:
            print("you are using my microphone again")
            voice = listener.listen(mic)
            request = listener.recognize_google(voice, language=lang)
            request = request.lower()
            print(request)
    except:
        request = 0
    print(request, "this is verifying")
    return request


def find_image(image):
    button_location = 0
    count = 0
    while not button_location:
        try:
            button_location = pyautogui.locateCenterOnScreen(f'{image}.png')
        except:
            button_location = 0
        if count > 250:
            break
        count = count + 1
    print(button_location)
    return button_location


def openApp():
    buttonx1, buttony1 = find_image("openapp")
    #buttonx1, buttony1 = find_image("https://i.ibb.co/jynrjsR/openapp2")
    pyautogui.moveTo(buttonx1, buttony1, 2, pyautogui.easeInOutQuad)
    pyautogui.click()


def whatsappMessage():
    speak("can you spell just the name")
    name = listen_to_verify(lang="tr-tr")
    #speak("can you spell just the surname")
    #surname = listen_to_verify(lang="tr-tr")
    speak(f"are you sure to send message to {name}, if no don't say anything")
    confirm = listen_to_verify()
    if str(confirm):
        msg = confirm.replace("yes", "")
        print(msg)
        if str(msg) != "0":
            speak("ok. I'm sending your message")
            pyautogui.press("win")
            time.sleep(4)
            pyautogui.write("whatsapp desktop")
            openApp()
            time.sleep(65)
            pyautogui.press("tab")
            time.sleep(1)
            pyautogui.write(name)
            time.sleep(2)
            pyautogui.press("tab")
            time.sleep(1)
            pyautogui.press("enter")
            time.sleep(1)
            pyautogui.write(msg, interval=0.24)
            pyautogui.press("enter")
            message_sent = 1
        else:
            speak("there was no message")
            message_sent = 0
            
    else:
        message_sent = 0
    return message_sent


def requests():
    rqst = listen()
    print(rqst)
    if rqst:
        print("here you are")
        print(rqst)
        if rqst.startswith(" can you" or " could you"):
            rqst = rqst.replace(" can you", "")
        if "play" in rqst:
            print("you did it")
            rqst = rqst.replace("play", "")
            if "the best song" in rqst:
                rqst = "deli vahit"
            pywhatkit.playonyt(rqst)
            rtrn = "OK."
        elif "stop" in rqst:
            pyautogui.press("playpause")
            rtrn = "I did it"
        elif "resume" in rqst:
            pyautogui.press("playpause")
            rtrn = "I believe you wanted this"
        elif "information about" in rqst:
            rqst = rqst.replace("information about", "")
            rtrn = wikipedia.summary(rqst, 2)
        elif "tell me a joke" in rqst:
            rtrn = pyjokes.get_joke()
        elif ("send message" in rqst) or ("send a message" in rqst):
            rqst = rqst.replace("send message", "")
            sent = whatsappMessage()
            if sent:
                rtrn = "your message sent successfully"
            else:
                rtrn = "message couldn't sent"
        elif "take note" in rqst:
            speak("what shall I note")
            note = listen_to_verify()
            pyautogui.press("win")
            time.sleep(4)
            pyautogui.write("Notepad", interval=0.15)
            openApp()
            time.sleep(4)
            pyautogui.write(note, interval=0.15)
            pyautogui.press("enter")
            rtrn = "Note taken."
        elif "open" in rqst:
            rqst = rqst.replace("open", "")
            pyautogui.press("win")
            time.sleep(4)
            pyautogui.write(rqst, interval=0.15)
            openApp()
            rtrn = "I opened it"
        else:
            try:
                rqst = rqst.replace("information about", "")
                rtrn = wikipedia.summary(rqst, 1)
                speak(f"searching for {rqst}")
                time.sleep(2)
                speak("result is: ")
            except:
                rtrn = "this request is not valid"
    else:
        rtrn = "Unfortunately you said nothing"
    speak(rtrn)


requests()  
