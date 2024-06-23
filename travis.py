import win32com.client
import speech_recognition as sr
import webbrowser
import random
import datetime


speak = win32com.client.Dispatch("SAPI.spvoice")


def take_command():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")

        try:
            audio = r.listen(source, timeout=5, phrase_time_limit=5)
        except sr.WaitTimeoutError:
            print("Listening timed out. Please try again.")
            return ""

    try:
        print("Recognizing...")
        # Use the Google Web Speech API
        r_query = r.recognize_google(audio)
        print(f"User said: {r_query}")
        return r_query
    except sr.UnknownValueError:
        print("Sorry, I didn't catch that. Can you please repeat?")
        return ""


if __name__ == "__main__":
    print("Enter the word you want to say:")
    s = input()
    speak.Speak(s)

    print("Listening...")
    query = take_command()

    sites = {"youtube": "https://www.youtube.com", "google": "https://www.google.com",
             "chat gpt": "https://chat.openai.com"}

    for site_key, site_url in sites.items():
        if f"open {site_key}".lower() in query.lower():
            speak.Speak(f"Opening {site_key.capitalize()}")
            webbrowser.open(site_url)
    a = "https://www.youtube.com/watch?v=bMD-NyruIBY"
    b = "https://www.youtube.com/watch?v=sWk9lpkGAfs"
    c = "://www.youtube.com/watch?v=N4GFIkcp5Wo"
    d = "https://www.youtube.com/watch?v=g5O5ufz8w34"
    music = [a, b, c, d]
    if "play music" in query.lower():
        speak.Speak("playing music")
        webbrowser.open(random.choice(music))
    if query:
        print("You said:", query)
        speak.Speak(query)
    if "tell time" in query.lower():
        strfTime = datetime.datetime.now().strftime(" %H hours and %M minutes")
        speak.Speak(f" it is {strfTime}")
