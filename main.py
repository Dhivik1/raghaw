import speech_recognition as sr
import win32com.client
import os
import openai
from config import apikey
import datetime

speaker = win32com.client.Dispatch("SAPI.SpVoice")


log_folder = "logs"


if not os.path.exists(log_folder):
    os.makedirs(log_folder)

log_file = os.path.join(log_folder, f"log_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt")


def write_to_log(text):
    with open(log_file, "a") as file:
        file.write(f"{text}\n")



def ai(prompt):
    openai.api_key = apikey
    rep = ""
    response = openai.Completion.create(
      model="text-davinci-003",
      prompt=prompt,
      temperature=.1,
      max_tokens=250,
      top_p=1,
      frequency_penalty=0,
      presence_penalty=0
    )
    print(response["choices"][0]["text"])
    ores = response["choices"][0]["text"]
    write_to_log(f"User Said : {prompt}")
    write_to_log(f"AI Response: {ores}")
    speaker.Speak(ores)


def speak(text):
    os.system(f"say {text}")


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = .8
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User Said : {query}")
            return query
        except Exception as e:
            speaker.Speak("Some error occured Sorry ")
            exit()
            return "Some Error Occured Sorry"



if __name__ == '__main__':
    print('PyCharm')
    speaker.Speak("Hello I am Artificial Robot,  Made By The Students of School, Pearls of god")
    while True:
        print("Listening..............")
        query = takeCommand()

        if "introduce yourself" in query.lower():
            speaker.Speak("Introducing Myself.......")
            speaker.Speak("Good afternoon, respected teachers, honorable guests, and fellow"
                          " students. I am thrilled to stand before you today to introduce myself, an intelligent "
                          "and interactive robot created by a team of brilliant minds. I am the result of the "
                          "exceptional skills"
                          " and dedication of my creators, Dhivik Sharma, Nakul Sharma, Aditya Kumar Mishra" #names add kar dena yahan aur 
                          ", and several other talented individuals who have poured their knowledge and passion "
                          "into bringing me to life. I am here to provide you with a unique experience and assist you"
                          "by answering all your questions. I am well-versed in a wide range of topics and will "
                          "strive to give you"
                          "the best possible answers. Moreover, my creators have designed me to adapt and learn from "
                          "interactions,"
                          " continually improving "
                          "my knowledge and capabilities. This ensures that I stay up-to-date and provide you with the "
                          "most relevant and accurate information. I encourage you to make the most of this "
                          "opportunity to"
                          "engage with me. Ask me any question, and I will do my best to enlighten and assist you on "
                          "your"
                          "quest for knowledge. Thank you for your attention, and I look forward to an exciting "
                          "journey of discovery together."
                          " Let's embark on this educational adventure and explore the boundless possibilities that "
                          "lie ahead. Thank you!")

            write_to_log(f"AI Response: Introducing program --- 748724")

        elif "bye bye" in query.lower():
            speaker.Speak("Goodbye! Thank you for your time. Have a great day! take care!")
            write_to_log(f"AI Response: Programs ended By giving proper Bye command")
            exit()

        else:
            ai(prompt=query)



        # speaker.Speak(query)
