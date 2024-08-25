import pyttsx3
import datetime
import speech_recognition as sr
import requests
import webbrowser as wb
import os
import pyautogui
import subprocess
import win32com.client
from transformers import pipeline

# Replace with your actual API keys
WEATHER_API_KEY = '83ed035d9eaf52b0a390e8f4cf71a179'
NEWS_API_KEY = 'b1e96668332142b2a9ea49ec70e4ca7a'

engine = pyttsx3.init()


def speak(audio) -> None:
    engine.say(audio)
    engine.runAndWait()


def time() -> None:
    Time = datetime.datetime.now().strftime("%I:%M:%S %p")
    speak("The current time is")
    speak(Time)
    print("The current time is", Time)


def date() -> None:
    day = datetime.datetime.now().day
    month = datetime.datetime.now().month
    year = datetime.datetime.now().year
    speak("The current date is")
    speak(f"{day} {month} {year}")
    print(f"The current date is {day}/{month}/{year}")


def wishme() -> None:
    print("Welcome back, sir!")
    speak("Welcome back, sir!")

    hour = datetime.datetime.now().hour
    if 4 <= hour < 12:
        speak("Good Morning, Sir!")
        print("Good Morning, Sir!")
    elif 12 <= hour < 16:
        speak("Good Afternoon, Sir!")
        print("Good Afternoon, Sir!")
    elif 16 <= hour < 24:
        speak("Good Evening, Sir!")
        print("Good Evening, Sir!")
    else:
        speak("Good Night, Sir. See you tomorrow")

    speak("Jarvis at your service, sir. Please tell me how may I help you.")
    print("Jarvis at your service, sir. Please tell me how may I help you.")


def screenshot() -> None:
    img = pyautogui.screenshot()
    img_path = os.path.expanduser("~\\Pictures\\ss.png")
    img.save(img_path)


def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language="en-in")
        print(f"You said: {query}")

    except Exception as e:
        print("Error:", e)
        speak("Please say that again")
        return "None"

    return query


def get_weather_by_coordinates(lat, lon):
    try:
        url = f"http://api.openweathermap.org/data/2.5/weather?lat={lat}&lon={lon}&appid={WEATHER_API_KEY}&units=metric"
        response = requests.get(url).json()

        if response.get("cod") == 200:
            weather = response["weather"][0]["description"]
            temperature = response["main"]["temp"]
            city = response.get("name", "Unknown location")
            speak(f"The weather in {city} is {weather} with a temperature of {temperature} degrees Celsius.")
            print(f"Weather in {city}: {weather}, {temperature}°C")
        else:
            error_message = response.get("message", "Unknown error occurred.")
            speak(f"Sorry, I couldn't fetch the weather information. Error: {error_message}")
            print(f"Weather info not found. Error: {error_message}")

    except requests.RequestException as e:
        speak("An error occurred while fetching the weather information.")
        print(f"Request error: {e}")
    except Exception as e:
        speak("An unexpected error occurred.")
        print(f"Unexpected error: {e}")


def search_news(query):
    try:
        url = f"https://newsapi.org/v2/everything?q={query}&apiKey={NEWS_API_KEY}"
        response = requests.get(url).json()

        if response.get("status") == "ok" and response.get("articles"):
            articles = response["articles"]
            for article in articles[:3]:
                title = article["title"]
                description = article["description"]
                speak(f"Here's a news article: {title}. {description}")
                print(f"Title: {title}\nDescription: {description}\n")
        else:
            speak("Sorry, I couldn't fetch the news articles.")
            print("No news articles found.")

    except requests.RequestException as e:
        speak("An error occurred while fetching the news.")
        print(f"Request error: {e}")
    except Exception as e:
        speak("An unexpected error occurred.")
        print(f"Unexpected error: {e}")


# Initialize the Hugging Face pipeline
generator = pipeline('text-generation', model='gpt2', max_length=50, pad_token_id=50256)  # Use GPT-2 model

chatStr = ""
speaker = win32com.client.Dispatch("SAPI.SpVoice")

# Predefined list of applications
apps = {
    "notepad": "notepad.exe",
    "calculator": "calc.exe",
    "paint": "mspaint.exe",
    "word": r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
    "excel": r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
    "powerpoint": r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
    "chrome": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    "firefox": r"C:\Program Files\Mozilla Firefox\firefox.exe",
    "edge": r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    "vscode": r"C:\Users\YourUsername\AppData\Local\Programs\Microsoft VS Code\Code.exe",
    "spotify": r"C:\Users\YourUsername\AppData\Roaming\Spotify\Spotify.exe",
    "vlc": r"C:\Program Files\VideoLAN\VLC\vlc.exe",
    "discord": r"C:\Users\YourUsername\AppData\Local\Discord\app-1.0.9003\Discord.exe",
    "skype": r"C:\Program Files (x86)\Microsoft\Skype for Desktop\Skype.exe",
    "zoom": r"C:\Users\YourUsername\AppData\Roaming\Zoom\bin\Zoom.exe",
    "onenote": r"C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE",
    "outlook": r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
    "photoshop": r"C:\Program Files\Adobe\Adobe Photoshop 2021\Photoshop.exe",
    "illustrator": r"C:\Program Files\Adobe\Adobe Illustrator 2021\Support Files\Contents\Windows\Illustrator.exe",
    "premiere": r"C:\Program Files\Adobe\Adobe Premiere Pro 2021\Adobe Premiere Pro.exe",
    "blender": r"C:\Program Files\Blender Foundation\Blender 2.93\blender.exe",
    "steam": r"C:\Program Files (x86)\Steam\steam.exe",
    "epic games": r"C:\Program Files (x86)\Epic Games\Launcher\Portal\Binaries\Win32\EpicGamesLauncher.exe",
    "battle.net": r"C:\Program Files (x86)\Battle.net\Battle.net Launcher.exe",
    "origin": r"C:\Program Files (x86)\Origin\Origin.exe",
    "uplay": r"C:\Program Files (x86)\Ubisoft\Ubisoft Game Launcher\Uplay.exe",
    "gimp": r"C:\Program Files\GIMP 2\bin\gimp-2.10.exe",
    "audacity": r"C:\Program Files (x86)\Audacity\audacity.exe",
    "filezilla": r"C:\Program Files\FileZilla FTP Client\filezilla.exe",
    "teamviewer": r"C:\Program Files (x86)\TeamViewer\TeamViewer.exe"
}

# Predefined list of websites and popular movies
sites = {
    "youtube": "https://www.youtube.com",
    "wikipedia": "https://www.wikipedia.com",
    "google": "https://www.google.com",
    "netflix": "https://www.netflix.com",
    "amazon prime": "https://www.primevideo.com",
    "disney plus": "https://www.disneyplus.com",
    "hulu": "https://www.hulu.com",
    "hbo max": "https://www.hbomax.com",
    "spotify": "https://www.spotify.com",
    "apple music": "https://www.music.apple.com",
    "pandora": "https://www.pandora.com",
    "soundcloud": "https://www.soundcloud.com",
    "imdb": "https://www.imdb.com",
    "rotten tomatoes": "https://www.rottentomatoes.com",
    "metacritic": "https://www.metacritic.com",
    "cnn": "https://www.cnn.com",
    "bbc": "https://www.bbc.com",
    "nytimes": "https://www.nytimes.com",
    "forbes": "https://www.forbes.com",
    "buzzfeed": "https://www.buzzfeed.com",
    "reddit": "https://www.reddit.com",
    "twitter": "https://www.twitter.com",
    "facebook": "https://www.facebook.com",
    "instagram": "https://www.instagram.com",
    "linkedin": "https://www.linkedin.com",
    "pinterest": "https://www.pinterest.com",
    "tumblr": "https://www.tumblr.com",
    "flickr": "https://www.flickr.com",
    "vimeo": "https://www.vimeo.com",
    "dailymotion": "https://www.dailymotion.com",
    "twitch": "https://www.twitch.tv"
}


def open_application(app_name):
    if app_name in apps:
        speak(f"Opening {app_name}")
        subprocess.Popen(apps[app_name])
    else:
        speak(f"Sorry, I don't know how to open {app_name}.")


def open_website(site_name):
    if site_name in sites:
        speak(f"Opening {site_name}")
        wb.open(sites[site_name])
    else:
        speak(f"Sorry, I don't know how to open {site_name}.")


def generate_response(prompt):
    responses = generator(prompt, max_length=100, num_return_sequences=1)
    response = responses[0]['generated_text']
    speak(response)
    print(f"AI response: {response}")


# Main loop
if __name__ == "__main__":
    wishme()
    while True:
        query = takecommand().lower()

        if "time" in query:
            time()

        elif "date" in query:
            date()

        elif "screenshot" in query:
            screenshot()
            speak("Screenshot taken")

        elif "open" in query:
            app_name = query.replace("open ", "").strip()
            open_application(app_name)

        elif "website" in query or "search" in query:
            site_name = query.replace("open website ", "").replace("search ", "").strip()
            open_website(site_name)

        elif "weather" in query:
            speak("Please provide the latitude and longitude.")
            try:
                lat = float(input("Enter latitude: "))
                lon = float(input("Enter longitude: "))
                get_weather_by_coordinates(lat, lon)
            except ValueError:
                speak("Invalid coordinates provided.")

        elif "news" in query:
            search_query = query.replace("news", "").strip()
            search_news(search_query)

        elif "quit" in query or "exit" in query:
            speak("Goodbye, sir. Have a great day!")
            break

        elif "generate" in query:
            prompt = query.replace("generate", "").strip()
            generate_response(prompt)

        else:
            speak("I did not understand that. Please say it again.")