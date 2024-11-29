import ctypes
import datetime
import platform
import sys
import time
import webbrowser
import pyautogui
import pyttsx3
import sbc
import speech_recognition as sr
import web_links  # Importing the web_links module
import requests
import wikipedia
import os
import psutil
import json
import win32clipboard

def initialize_engine():
    """Initialize the text-to-speech engine."""
    engine = pyttsx3.init("sapi5")
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[2].id)  # Use a female voice
    engine.setProperty('rate', 150)  # Set speech rate
    engine.setProperty('volume', 1.0)  # Set volume
    return engine

def speak(text):
    """Convert text to speech."""
    engine = initialize_engine()
    engine.say(text)
    engine.runAndWait()

def listen():
    """Capture and recognize speech input."""
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source, duration=0.5)
        print("Listening...")
        try:
            audio = recognizer.listen(source, timeout=10, phrase_time_limit=5)
            print("Recognizing...")
            command = recognizer.recognize_google(audio, language='en-in')
            print(f"You said: {command}")
            return command.lower()
        except sr.WaitTimeoutError:
            pass
        except sr.UnknownValueError:
            pass
        except sr.RequestError as e:
            print(f"Speech recognition service error: {e}")
            speak("There seems to be an issue with the speech recognition service.")
        return None

def load_applications(file_path="applications.json"):
    """Load application paths from a JSON file."""
    try:
        with open(file_path, "r") as file:
            applications = json.load(file)
        return applications
    except FileNotFoundError:
        print("Application file not found. Ensure 'applications.json' is in the project directory.")
        return {}
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON: {e}")
        return {}

def open_application(command, applications):
    """Open an application based on the command."""
    command = command.lower()
    if command in applications:
        app_path = applications[command]
        if os.path.exists(app_path):
            os.startfile(app_path)  # Open the application
            speak(f"Opening {command}.")
            return True
        else:
            speak(f"The path for {command} is invalid or the program is not installed.")
            return False
    return False  # Command not found in the applications file


# Function to normalize the text by converting to lowercase and handling apostrophes properly
def normalize_text(text):
    """Normalize text by converting to lowercase and handling apostrophes."""
    # Convert text to lowercase
    text = text.lower()

    # Return the normalized text, keeping apostrophes intact
    return text


# Function to load responses from the JSON file and normalize keys
def load_responses(file_path="responses.json"):
    """Load custom responses from a JSON file and normalize keys."""
    try:
        with open(file_path, "r") as file:
            responses = json.load(file)

        # Normalize all keys (convert to lowercase while keeping apostrophes intact)
        responses = {normalize_text(key): value for key, value in responses.items()}
        return responses
    except FileNotFoundError:
        print("Responses file not found. Ensure 'responses.json' is in the project directory.")
        return {}
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON: {e}")
        return {}


# Function to handle the conversation and respond based on user input
def handle_conversation(command, responses):
    """Handle custom conversational commands using loaded responses."""
    # Normalize the command (convert to lowercase while keeping apostrophes intact)
    normalized_command = normalize_text(command)

    # Check if the normalized command exists in the normalized responses
    return responses.get(normalized_command, "Sorry, I didn't understand that.")  # Default response if no match


# Function to update or add a new response to the JSON file
def update_responses(command, response, file_path="responses.json"):
    """Update or add a new response to the JSON file."""
    try:
        # Load existing responses
        responses = load_responses(file_path)

        # Update or add the new response with normalized command key
        responses[normalize_text(command)] = response

        # Save back to the JSON file
        with open(file_path, "w") as file:
            json.dump(responses, file, indent=4)

        print(f"The response for '{command}' has been updated.")
    except Exception as e:
        print(f"Error updating responses: {e}")
        print("I couldn't update the responses. Please try again.")

def handle_command(command):
    """Process the voice command."""
    print(f"Received command: {command}")

    # Load applications from the JSON file
    applications = load_applications()

    # Check if the command matches an application
    if open_application(command, applications):
        return True

    # Handle other commands like custom responses
    responses = load_responses()
    response = handle_conversation(command, responses)
    if response:
        speak(response)
        return True

    # Handle additional logic (e.g., web links, system status)
    if "open" in command:
        if web_links.handle_web_link_command(command):
            return True

    return False  # Command wasn't handled successfully
def update_application(command, app_path, file_path="applications.json"):
    """Update or add a new application path to the JSON file."""
    try:
        # Load existing applications
        applications = load_applications(file_path)

        # Update or add the new application path
        applications[command.lower()] = app_path

        # Save back to the JSON file
        with open(file_path, "w") as file:
            json.dump(applications, file, indent=4)
        speak(f"The application path for '{command}' has been updated.")
    except Exception as e:
        print(f"Error updating applications: {e}")
        speak("I couldn't update the application. Please try again.")

def cal_day():
    day = datetime.datetime.today().weekday() + 1
    day_dict={
        1:"Monday",
        2:"Tuesday",
        3:"Wednesday",
        4:"Thursday",
        5:"Friday",
        6:"Saturday",
        7:"Sunday"
    }
    if day in day_dict.keys():
        day_of_week = day_dict[day]
        print(day_of_week)
    return day_of_week

def wishMe():
    hour = int(datetime.datetime.now().hour)
    t = time.strftime("%I:%M:%p")
    day = cal_day()

    if(hour>=0) and (hour<=12) and ('AM' in t):
        speak(f"Good morning Sir, How can I assist you?")
    elif(hour>=12)  and (hour<=16) and ('PM' in t):
        speak(f"Good afternoon Sir. How can I assist you?")
    else:
        speak(f"Good evening Sir, How can I help you?")

def fetch_latest_news():
    try:
        url = "https://newsapi.org/v2/top-headlines"
        params = {
            "country": "us",  # Replace with your preferred country code
            "apiKey": "2434336a8fa84692896eb42b2dcf586b",  # Replace with your NewsAPI key
            "pageSize": 10
        }
        response = requests.get(url, params=params)
        news_data = response.json()

        if news_data["status"] == "ok":
            articles = news_data["articles"]
            if articles:
                speak("Here are the latest top news headlines.")
                for i, article in enumerate(articles):
                    headline = article["title"]
                    print(f"Headline {i + 1}: {headline}")
                    speak(f"Headline {i + 1}: {headline}")
            else:
                speak("I couldn't find any news at the moment.")
        else:
            speak("There was a problem fetching the news.")
    except Exception as e:
        print(f"Error fetching news: {e}")
        speak("I encountered an issue while fetching the news.")

def fetch_topic_summary(topic):
    """Fetch a brief summary of the specified topic from Wikipedia."""
    try:
        summary = wikipedia.summary(topic, sentences=3)  # Get a short summary of the topic
        print(f"Summary of {topic}: {summary}")
        speak(f"Here is what I found about {topic}: {summary}")
    except wikipedia.exceptions.DisambiguationError as e:
        print(f"Disambiguation error: {e}")
        speak(f"The topic {topic} is too broad. Can you be more specific?")
    except wikipedia.exceptions.PageError:
        print(f"No Wikipedia page found for {topic}.")
        speak(f"I couldn't find anything about {topic}. Please try a different topic.")
    except Exception as e:
        print(f"Error fetching summary: {e}")
        speak("I encountered an issue while fetching information about the topic.")

def get_current_time(command):
    """Speak the current time."""
    if "current time" in command or "what is the time" in command:
        # Get the current time
        now = datetime.datetime.now()
        current_time = now.strftime("%I:%M %p")  # Format: HH:MM AM/PM

        # Respond with the current time
        speak(f"The current time is {current_time}.")
        print(f"Time announced: {current_time}")

def switch_to_previous_window(command):
    """Switch to the previous window."""
    if "previous window" in command or "go back to previous window" in command:
        pyautogui.hotkey("alt", "shift", "tab")  # Simulate Alt + Shift + Tab
        speak("Switched to the previous window.")
        print("Switched to the previous window.")

def switch_to_next_window(command):
    """Switch to the next window."""
    if "next window" in command or "switch to next window" in command:
        pyautogui.hotkey("alt", "tab")  # Simulate Alt + Tab
        speak("Switched to the next window.")
        print("Switched to the next window.")

def get_current_date(command):
    """Speak the current day, date, and year."""
    if "what day is today" in command or "tell me the date" in command or "today's date" in command:
        # Get the current day, date, and year
        now = datetime.datetime.now()
        day = now.strftime("%A")  # e.g., "Monday"
        date = now.strftime("%B %d, %Y")  # e.g., "November 28, 2024"

        # Construct the response
        response = f"Today is {day}, {date}."

        # Speak and print the response
        speak(response)
        print(f"Date announced: {response}")

def browsing(query):
    """Search on Google or YouTube or open Chrome."""
    # Check for Google search
    if 'google' in query:
        speak("What should I search on Google?")
        search_query = listen()  # Get the search input from the user
        if search_query:  # Ensure the user provides input
            # Construct the proper Google search URL
            search_url = f"https://www.google.com/search?q={search_query.replace(' ', '+')}"
            webbrowser.open(search_url)  # Open the browser with the search URL
            speak(f"Here are the search results for {search_query}.")
        else:
            speak("I didn't catch that. Please try again.")

    # Check for YouTube search
    elif "check on youtube" in query or "search youtube" in query:
        speak("What should I search on YouTube?")
        search_query = listen()  # Get the search input from the user
        if search_query:  # Ensure the user provides input
            # Construct the proper YouTube search URL
            search_url = f"https://www.youtube.com/results?search_query={search_query.replace(' ', '+')}"
            webbrowser.open(search_url)  # Open the browser with the YouTube search URL
            speak(f"Here are the search results for {search_query} on YouTube.")
        else:
            speak("I didn't catch that. Please try again.")

def search_files(query, search_type):
    """Search for files by type (video, audio, document, image)"""
    file_extensions = {
        "video": [".mp4", ".mkv", ".avi", ".mov"],
        "audio": [".mp3", ".wav", ".aac"],
        "document": [".pdf", ".docx", ".txt", ".pptx", ".xlsx"],
        "image": [".jpg", ".jpeg", ".png", ".bmp", ".webp"]
    }
    if search_type not in file_extensions:
        speak("Invalid search type. Please specify video, audio, or document.")
        return

    base_path = "D:\\"  # Change to the desired starting directory
    results = []

    for root, dirs, files in os.walk(base_path):
        for file in files:
            if any(file.endswith(ext) for ext in file_extensions[search_type]) and query.lower() in file.lower():
                results.append(os.path.join(root, file))

    if results:
        speak(f"I found {len(results)} matching {search_type} files.")
        print("Here are the matching files:")
        for result in results:  # Display all results in the console
            print(result)
    else:
        speak(f"No {search_type} files found matching your query.")

def condition():
    """Check system status: CPU usage and battery percentage."""
    usage = str(psutil.cpu_percent())
    speak(f"CPU is at {usage} percentage")
    battery = psutil.sensors_battery()
    percentage = battery.percent
    speak(f"Boss, our system has {percentage} percentage battery.")

    if percentage >= 80:
        speak("Boss, we could have enough charging to continue our recording.")
    elif percentage >= 40 and percentage <= 75:
        speak("Boss, we should connect our system to a charging point to charge our battery.")
    else:
        speak("Boss, we have very low power. Please connect to charging, otherwise recording should be off.")

def open_new_tab(command):
    """Open a new tab in the default browser."""
    if "new tab" in command or "open new tab" in command:
        webbrowser.open_new_tab("about:blank")  # Opens a blank new tab
        speak("A new tab has been opened.")
        print("New tab opened.")

def close_browser_tab():
    """Close the browser tab using platform-specific hotkeys."""
    try:
        system = platform.system()
        if system == "Windows" or system == "Linux":
            pyautogui.hotkey("ctrl", "w")  # Close tab for Windows/Linux
        elif system == "Darwin":  # macOS
            pyautogui.hotkey("command", "w")  # Close tab for macOS
        speak("")
    except Exception as e:
        print(f"Error while trying to close the tab: {e}")
        speak("I encountered an issue while trying to close the tab.")

def open_power_options(command):
    """Open Power Options."""
    if "power options" in command:
        os.system("powercfg.cpl")
        speak("Opening Power Options.")
        print("Power Options opened.")

def open_device_manager(command):
    """Open Device Manager."""
    if "device manager" in command:
        os.system("devmgmt.msc")
        speak("Opening Device Manager.")
        print("Device Manager opened.")

def open_sounds(command):
    """Open Sound settings."""
    if "sound settings" in command or "sounds" in command:
        os.system("mmsys.cpl sounds")
        speak("Opening Sound settings.")
        print("Sound settings opened.")

def open_system_settings(command):
    """Open System settings."""
    if "system settings" in command or "system" in command:
        os.system("sysdm.cpl")
        speak("Opening System settings.")
        print("System settings opened.")

def show_my_computer(command):
    """Open This PC (My Computer)."""
    if "my computer" in command or "this pc" in command:
        os.system("explorer shell:MyComputerFolder")
        speak("Opening This PC.")
        print("This PC opened.")

def open_task_manager(command):
    """Open Task Manager."""
    if "task manager" in command:
        os.system("start taskmgr")
        speak("Opening Task Manager.")
        print("Task Manager opened.")

if __name__ == "__main__":
    wishMe()
    while True:
        command = listen()  # Listen for a command
        if command:
            # Check for commands and handle them accordingly
            if "exit" in command or "quit" in command or "stop" in command:
                speak("Goodbye!")
                break
            elif "tell me about" in command or "about" in command:
                topic = command.replace("tell me about", "").replace("about", "").strip()
                if topic:
                    fetch_topic_summary(topic)  # Fetch the topic summary
                else:
                    speak("I didn't understand the topic. Could you please say the topic again?")
            elif "wikipedia" in command:
                topic = command.replace("wikipedia", "").strip()
                if topic:
                    fetch_topic_summary(topic)
                else:
                    speak("What topic should I search for on Wikipedia?")
            elif "latest news" in command or "top news" in command:
                fetch_latest_news()
            elif "google" in command or "check on google" in command or "open chrome" in command or "youtube" in command:
                browsing(command)  # Handle browsing-related commands
            elif "search video" in command:
                speak("What video file should I look for?")
                search_query = listen()  # Get the search query from the user
                if search_query:
                    search_files(search_query, "video")  # Search for video files
            elif "search audio" in command:
                speak("What audio file should I look for?")
                search_query = listen()  # Get the search query from the user
                if search_query:
                    search_files(search_query, "audio")  # Search for audio files
            elif "search document" in command:
                speak("What document should I look for?")
                search_query = listen()  # Get the search query from the user
                if search_query:
                    search_files(search_query, "document")  # Search for document files
            elif "search image" in command:
                speak("What image file should I look for?")
                search_query = listen()  # Get the search query from the user
                if search_query:
                    search_files(search_query, "image")  # Search for image files
            elif "system condition" in command or "condition of the system" in command:
                speak("Checking the system condition.")
                condition()  # Check system status
            elif "close tab" in command:
                close_browser_tab()  # Close the browser tab
            elif "new tab" in command or "open new tab" in command:
                open_new_tab(command)
            elif ("volume up" in command) or ("increase volume" in command):
                pyautogui.press("volumeup")
                speak("Volume increased")
            elif ("volume down" in command) or ("decrease volume" in command):
                pyautogui.press("volumedown")
                speak("Volume decreased")
            elif ("volume mute" in command) or ("mute the sound" in command):
                pyautogui.press("volumemute")
                speak("Volume muted")
            elif ("take screenshot" in command) or ("screenshot" in command):
                # Simulate pressinwog the "Print Screen" key
                pyautogui.press("printscreen")
                speak("Screenshot taken")
            elif "current time" in command or "what is the time" in command or "time please" in command or "present time" in command:
                get_current_time(command)
            elif "what day is today" in command or "tell me the date" in command or "today's date" in command:
                get_current_date(command)
            elif ("zoom in" in command):
                # Simulate pressing "Ctrl" and "+" together for zooming in
                pyautogui.hotkey("ctrl", "+")
                speak("Zooming in")
            elif ("zoom out" in command):
                # Simulate pressing "Ctrl" and "-" together for zooming out
                pyautogui.hotkey("ctrl", "-")
                speak("Zooming out")
            elif ("scroll down" in command):
                # Simulate scrolling down (negative value scrolls down)
                pyautogui.scroll(-1000)  # Adjust the scroll amount as needed
                speak("Scrolling down")
            elif ("scroll up" in command):
                # Simulate scrolling up (positive value scrolls up)
                pyautogui.scroll(1000)  # Adjust the scroll amount as needed
                speak("Scrolling up")
            elif ("scroll at the bottom" in command):
                # Scroll all the way to the bottom by scrolling down a large amount
                pyautogui.scroll(-10000)  # A large negative value to simulate scrolling to the bottom
                speak("Scrolling to the bottom")
            elif ("scroll at the top" in command):
                # Scroll all the way to the top by scrolling up a large amount
                pyautogui.scroll(10000)  # A large positive value to simulate scrolling to the top
                speak("Scrolling to the top")
            elif ("minimise window" in command or "minimize window" in command):
                # Simulate pressing "Alt + Space" then "N" to minimize the window
                pyautogui.hotkey("alt", "space")  # Open the window's system menu
                pyautogui.press("n")  # Select the "Minimize" option
                speak("Window minimized")
            elif ("maximize window" in command or "maximise window" in command):
                # Simulate pressing "Alt + Space" then "X" to maximize the window
                pyautogui.hotkey("alt", "space")  # Open the window's system menu
                pyautogui.press("x")  # Select the "Maximize" option
                speak("Window maximized")
            elif "close window" in command:
                # Simulate pressing "Alt + F4" to close the current window
                pyautogui.hotkey("alt", "f4")
                speak("Window closed")
            elif "lock screen" in command:
                ctypes.windll.user32.LockWorkStation()  # Lock the system screen
                speak("Screen locked.")
            elif "start menu" in command:
                pyautogui.hotkey("win")  # Simulate pressing the Windows key
                speak("Start menu opened.")
            elif "control panel" in command:
                os.system("start control")  # Open the Control Panel
                speak("Control panel opened.")
            elif "switch to next window" in command:
                pyautogui.hotkey("alt", "tab")  # Simulate pressing Alt + Tab
                speak("Switched to next window.")
            elif "toggle full screen" in command:
                pyautogui.hotkey("alt", "enter")  # Toggle full screen in many apps
                speak("Toggled full-screen mode.")
            elif "enable bluetooth" in command:
                os.system("start ms-settings:bluetooth")  # Opens Bluetooth settings on Windows
                speak("Bluetooth enabled.")
            elif "disable bluetooth" in command:
                os.system(
                    "start ms-settings:bluetooth")  # Same as enabling; it can open Bluetooth settings for manual toggling
                speak("Bluetooth disabled.")
            elif "previous window" in command or "go back to previous window" in command:
                switch_to_previous_window(command)
            elif "next window" in command or "switch to next window" in command:
                switch_to_next_window(command)
            elif "power options" in command:
                open_power_options(command)
            elif "device manager" in command:
                open_device_manager(command)
            elif "sound settings" in command or "sounds" in command:
                open_sounds(command)
            elif "system settings" in command or "system" in command:
                open_system_settings(command)
            elif "my computer" in command or "this pc" in command or "show my computer" in command:
                show_my_computer(command)
            elif "task manager" in command:
                open_task_manager(command)
            elif "exit" in command:
                sys.exit()
            else:
                # Handle opening websites through voice commands
                if not handle_command(command):
                    speak(f"I heard: {command}, but I don't recognize that as a valid request.")
