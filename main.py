import speech_recognition as sr
import os
import win32com.client


speaker = win32com.client.Dispatch("SAPI.SpVoice")


def speak(text):
    os.system(f"say {text}")


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = .09
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User Said : {query}")
            return query
        except Exception as e:
            print("Some error occurred coud'nt understand your voice",str(e))
            return "I apologize ! Some error Occurred from My side"


if __name__ == '__main__':
    print('PyCharm')
    speaker.Speak("Hello I am Nakul AI")
    while True:
        print("Listening..............")
        query = takeCommand()
        if "Introduce Yourself ".lower() in query.lower():
            speaker.Speak("Introducing Myself.......")
            speaker.Speak("Good afternoon, respected teachers, honorable guests, and fellow"
                          " students. I am thrilled to stand before you today to introduce myself, an intelligent "
                          "and interactive robot created by a team of brilliant minds. My name is nakul, a well known,"
                          "dog , just kidding ,"
                          " and I am the result of the exceptional skills"
                          " and dedication of my creators, Dhivik Sharma, Nakul Sharma, Aditya Kumar Mishra"
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
                          "quest for knowledge.Thank you for your attention, and I look forward to an exciting "
                          "journey of discovery together."
                          "Let's embark on this educational adventure and explore the boundless possibilities that "
                          "lie ahead. Thank you!")

        if "Bye Bye".lower() in query.lower():
            speaker.Speak("Goodbye! Thank you for your time. Have a great day! take care!")
            exit()
        # speaker.Speak(query)
