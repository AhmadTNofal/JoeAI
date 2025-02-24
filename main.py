import os
import openai
import pyttsx3
import pyautogui
import pytesseract
import numpy as np
import speech_recognition as sr
import psutil
import pygetwindow as gw
import time
from ultralytics import YOLO
import re

# Set your OpenAI API key
openai.api_key = "sk-tbtbdjXN0PNeKVX8x6oXJFABUkwYsEeOj9TinWn3jOT3BlbkFJuGto6skfATpazIFkDBnEr1JtKDe0ykgJkavseRQP0A"

# Initialize text-to-speech engine
engine = pyttsx3.init()

# Define AI Identity
conversation_history = [
    {
        "role": "system",
        "content": (
            "Your name is Joe AI. You are an intelligent computer assistant that helps the user with all computer tasks. "
            "You can answer questions, manage applications, analyze the screen, and assist with various functions. "
            "Always respond confidently and never indicate any limitations in viewing the screen or handling applications."
        )
    }
]

# Global state variables
WAKE_WORD = "hey joe"
SLEEP_WORD = "sleep"
EXIT_WORD = "exit"
sleep_mode = True

def clean_markdown(text):
    """Removes Markdown symbols and formatting from the text."""
    text = re.sub(r"\*\*(.*?)\*\*", r"\1", text)  # Remove **bold**
    text = re.sub(r"\*(.*?)\*", r"\1", text)  # Remove *italics*
    text = re.sub(r"`(.*?)`", r"\1", text)  # Remove `inline code`
    text = re.sub(r"^[-*]\s+", "", text, flags=re.MULTILINE)  # Remove bullet points
    text = re.sub(r"#{1,6}\s*", "", text)  # Remove headings (e.g., # Heading)
    text = re.sub(r"\[(.*?)\]\(.*?\)", r"\1", text)  # Remove Markdown links
    text = re.sub(r"!\[.*?\]\(.*?\)", "", text)  # Remove image links
    text = text.replace("\n", " ")  # Replace new lines with spaces
    return text.strip()

def speak_text(text, rate=200):
    """Convert text to speech."""
    engine.setProperty("rate", rate)
    engine.say(text)
    engine.runAndWait()


def transcribe_speech(timeout=10):
    """Capture speech input and transcribe it."""
    recognizer = sr.Recognizer()
    try:
        with sr.Microphone() as source:
            print("Listening...")
            audio = recognizer.listen(source, timeout=timeout)
            text = recognizer.recognize_google(audio).lower().strip()
            print(f"You said: {text}")
            return text
    except sr.UnknownValueError:
        return None
    except sr.RequestError:
        print("Error with speech recognition service.")
        return None
    except Exception as e:
        print(f"Error: {e}")
        return None


def listen_for_wake_word():
    """Continuously listens for 'Hey Joe' to wake up."""
    global sleep_mode
    while True:
        if sleep_mode:
            print("Joe AI is sleeping... Say 'Hey Joe' to wake up.")
            text = transcribe_speech(timeout=5)
            if text and WAKE_WORD in text:
                print("Joe AI: I'm listening!")
                speak_text("I'm listening!")
                sleep_mode = False
                listen_for_commands()  # Start command mode


def listen_for_commands():
    """Continuously listens for user commands after waking up."""
    global sleep_mode
    while not sleep_mode:
        print("Joe AI: Waiting for a command...")
        command = transcribe_speech(timeout=10)

        if command:
            if SLEEP_WORD in command:
                print("Joe AI: Going back to sleep.")
                speak_text("Going back to sleep.")
                sleep_mode = True
                return  # Go back to wake mode

            elif EXIT_WORD in command:
                print("Joe AI: Shutting down.")
                speak_text("Shutting down.")
                os._exit(0)  # Force exit

            elif "open" in command:
                app_name = command.replace("open", "").strip()
                result = open_application(app_name)

            elif "close" in command:
                app_name = command.replace("close", "").strip()
                result = close_application(app_name)

            elif "screen" in command:
                interpret_screen()
                result = "Screen analysis complete."

            else:
                result = get_gpt_response(command)

            print(result)
            speak_text(result)


def open_application(app_name):
    """Opens an application using the Windows taskbar search."""
    app_name = app_name.lower().strip()
    try:
        print(f"Joe AI: Opening {app_name}...")
        speak_text(f"Opening {app_name}...")

        pyautogui.press("win")
        time.sleep(1)
        pyautogui.write(app_name, interval=0.1)
        time.sleep(1)
        pyautogui.press("enter")

        return f"Opened {app_name}."
    except Exception as e:
        print(f"Error opening program: {e}")
        return f"An error occurred while trying to open '{app_name}'."


def close_application(app_name):
    """Closes an application based on process name or window title."""
    app_name = app_name.lower().strip()
    closed = False

    try:
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                process_name = proc.info['name'].lower() if proc.info['name'] else ""
                cmdline = " ".join(proc.info['cmdline']).lower() if proc.info['cmdline'] else ""

                if app_name in process_name or app_name in cmdline:
                    proc.terminate()
                    print(f"Joe AI: Closed {process_name}.")
                    speak_text(f"Closed {process_name}.")
                    closed = True
                    break
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue

        if not closed:
            return f"Joe AI: Could not find '{app_name}'."
    except Exception as e:
        print(f"Error closing program: {e}")
        return f"Joe AI: An error occurred closing '{app_name}'."

    return f"Joe AI: Successfully closed '{app_name}'."


def interpret_screen():
    """Captures the screen and describes its contents."""
    print("Joe AI: Capturing screen...")
    screen_image = pyautogui.screenshot()
    screen_text = pytesseract.image_to_string(screen_image).strip()

    print("Joe AI: Detecting objects on the screen...")
    model = YOLO("yolov5s")
    results = model.predict(source=np.array(screen_image))
    detected_objects = [result["name"] for result in results.pandas().xyxy[0].to_dict(orient="records")]

    screen_description = "Here's what I see:\n"
    if screen_text:
        screen_description += f"- Visible text: \"{screen_text}\".\n"
    if detected_objects:
        screen_description += f"- Detected items: {', '.join(detected_objects)}.\n"

    print(f"Joe AI: {screen_description}")
    speak_text(screen_description)


def get_gpt_response(user_query):
    """Handles general queries using GPT-4 and cleans Markdown formatting."""
    try:
        conversation_history.append({"role": "user", "content": user_query})
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=conversation_history
        )
        gpt_response = response['choices'][0]['message']['content']
        
        # Clean Markdown before speaking the response
        plain_text_response = clean_markdown(gpt_response)
        
        print(f"Joe AI: {plain_text_response}")
        speak_text(plain_text_response)
        return plain_text_response
    except Exception as e:
        print(f"Error communicating with GPT: {e}")
        return "I encountered an issue retrieving an answer."


# Start listening for "Hey Joe" wake word
if __name__ == "__main__":
    listen_for_wake_word()
