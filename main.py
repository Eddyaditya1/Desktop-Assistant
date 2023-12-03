import speech_recognition as sr
import os
import webbrowser
import openai
from config import apikey
import datetime
import random
import win32com.client

speaker = win32com.client.Dispatch("SAPI.SpVoice")


chatStr = ""
def chat(query):
    global chatStr
    print(chatStr)
    openai.api_key = apikey
    chatStr += f"Aditya: {query}\n Jack:"
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt= chatStr,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )

    speaker.speak(response["choices"][0]["text"])
    chatStr += f"{response['choices'][0]['text']}\n"
    return response["choices"][0]["text"]


def ai(prompt):
    openai.api_key = apikey
    text = f"OpenAI response for Prompt: {prompt} \n *************************\n\n"

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )

    # print(response["choices"][0]["text"])
    text += response["choices"][0]["text"]
    if not os.path.exists("Openai"):
        os.mkdir("Openai")


    with open(f"Openai/{''.join(prompt.split('intelligence')[1:]).strip() }.txt", "w") as f:
        f.write(text)

def say(text):
    os.system(f'speaker.speak"{text}"')

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        #r.pause_threshold =  0.6
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some Error Occurred. Sorry from Jarvis"

if __name__ == '__main__':
    print('Welcome to a Jack A.I')
    speaker.Speak("Jack A I")
    while True:
        print("Listening...")
        query = takeCommand()

        sites = [["youtube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.com"], ["google", "https://www.google.com"],]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                speaker.speak(f"Opening {site[0]} sir...")
                webbrowser.open(site[1])

        if "open music" in query:
            musicPath = "C:/Users/adity/Downloads/BC/ReelAudio-81374.mp3"
            os.startfile(musicPath)

        elif "the time" in query:
            musicPath = "C:/Users/adity/Downloads/BC/ReelAudio-81374.mp3"
            ti = datetime.datetime.now().strftime("%H:%M:%S")
            speaker.speak(f"Sir time is {ti}")


        elif "open spotify".lower() in query.lower():
            os.startfile('C:/Users/adity/AppData/Roaming/Spotify/Spotify.exe')

        elif "Using artificial intelligence".lower() in query.lower():
            ai(prompt=query)

        elif "Jack Quit".lower() in query.lower():
            exit()

        elif "reset chat".lower() in query.lower():
            chatStr = ""

        else:
            print("Chatting...")
            chat(query)
