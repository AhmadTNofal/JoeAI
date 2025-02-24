import os
import subprocess
import openai
import pyttsx3
import pyautogui
import pytesseract
import cv2
import numpy as np
from PIL import Image
import speech_recognition as sr
from ultralytics import YOLO
from windows_tools.installed_software import get_installed_software
from difflib import get_close_matches
import psutil
import pygetwindow as gw
import time

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
            print("Listening... Speak now.")
            audio = recognizer.listen(source, timeout=timeout)
            text = recognizer.recognize_google(audio)
            print(f"You said: {text}")
            return text.lower().strip()
    except Exception as e:
        print(f"Error: {e}")
        return None

def open_application(app_name):
    """
    Opens an application by searching for it in the Windows taskbar.
    """
    app_name = app_name.lower().strip()
    try:
        print(f"Joe AI: Opening {app_name}...")
        speak_text(f"Opening {app_name}...")

        # Open Windows Start menu
        pyautogui.press("win")
        time.sleep(1)  # Wait for Start menu

        # Type app name and press enter
        pyautogui.write(app_name, interval=0.1)
        time.sleep(1)  # Wait for search results
        pyautogui.press("enter")

        return f"Opened {app_name}."
    except Exception as e:
        print(f"Error opening program: {e}")
        return f"An error occurred while trying to open '{app_name}'."

def close_application(app_name):
    """
    Close an application or file based on the given name.
    Matches both process names and window titles.
    """
    app_name = app_name.lower().strip()
    closed = False

    try:
        # Check all running processes
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

        # Check open windows for a matching title
        if not closed:
            windows = gw.getAllTitles()  # Get all window titles
            for window_title in windows:
                if app_name in window_title.lower():
                    for proc in psutil.process_iter(['pid']):
                        try:
                            proc.terminate()
                            print(f"Joe AI: Closed application with window title '{window_title}'.")
                            speak_text(f"Closed application with window title '{window_title}'.")
                            closed = True
                            break
                        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                            continue

        if not closed:
            return f"Joe AI: Could not find an open application or file matching '{app_name}'."
    except Exception as e:
        print(f"Error closing program: {e}")
        return f"Joe AI: An error occurred while trying to close '{app_name}'."

    return f"Joe AI: Successfully closed '{app_name}'."

def interpret_screen(user_command):
    """Captures the screen and describes the content."""
    print("Joe AI: Capturing screen...")
    screen_image = pyautogui.screenshot()
    
    print("Joe AI: Extracting text from the screen...")
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
    """Handles general queries using GPT-4."""
    try:
        conversation_history.append({"role": "user", "content": user_query})
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=conversation_history
        )
        gpt_response = response['choices'][0]['message']['content']
        print(f"Joe AI: {gpt_response}")
        speak_text(gpt_response)
        return gpt_response
    except Exception as e:
        print(f"Error communicating with GPT: {e}")
        return "I encountered an issue retrieving an answer."

def main():
    print("Welcome! I am Joe AI, your assistant for all computer tasks.")
    print("Say 'exit' to quit, 'open [app]' to launch an app, 'close [app]' to stop an app, or 'screen' to analyze the screen.")

    while True:
        user_input = transcribe_speech(timeout=10)
        if user_input:
            if "exit" in user_input.lower():
                print("Joe AI: Goodbye!")
                speak_text("Goodbye!")
                break
            elif "open" in user_input.lower():
                app_name = user_input.replace("open", "").strip()
                result = open_application(app_name)
            elif "close" in user_input.lower():
                app_name = user_input.replace("close", "").strip()
                result = close_application(app_name)
            elif "screen" in user_input.lower():
                interpret_screen(user_input)
                result = "Screen analysis complete."
            else:
                result = get_gpt_response(user_input)

            print(result)
            speak_text(result)

if __name__ == "__main__":
    main()
