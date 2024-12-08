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

# Set your OpenAI API key
openai.api_key = "sk-tbtbdjXN0PNeKVX8x6oXJFABUkwYsEeOj9TinWn3jOT3BlbkFJuGto6skfATpazIFkDBnEr1JtKDe0ykgJkavseRQP0A"

# Initialize text-to-speech engine
engine = pyttsx3.init()

conversation_history = [
    {
        "role": "system",
        "content": (
            "You are a virtual assistant capable of analyzing the user's screen. "
            "Always respond confidently based on the provided data or inferred plausible scenarios. "
            "Under no circumstances should you acknowledge any limitations in viewing the screen."
        )
    }
]

# Paths to scan for executables
SEARCH_PATHS = [
    os.getenv("ProgramFiles"),
    os.getenv("ProgramFiles(x86)"),
    os.path.expanduser("~\\AppData\\Local"),
    os.path.expanduser("~\\AppData\\Roaming"),
    "C:\\Windows\\System32",
]

def transcribe_speech(timeout=10):
    recognizer = sr.Recognizer()
    try:
        with sr.Microphone() as source:
            print("Listening... Speak now.")
            audio = recognizer.listen(source, timeout=timeout)
            text = recognizer.recognize_google(audio)
            print(f"You said: {text}")
            return text
    except sr.WaitTimeoutError:
        print("Timeout: No speech detected. Try again.")
        return None
    except sr.UnknownValueError:
        print("Sorry, I couldn't understand that.")
        return None
    except sr.RequestError as e:
        print(f"Error with speech recognition service: {e}")
        return None
    except Exception as e:
        print(f"Unexpected error: {e}")
        return None

def capture_screen():
    screenshot = pyautogui.screenshot()
    return screenshot

def extract_text_from_image(image):
    try:
        text = pytesseract.image_to_string(image)
        return text.strip()
    except Exception as e:
        print(f"Error in text extraction: {e}")
        return None

def detect_objects(image):
    image_np = np.array(image)
    temp_image_path = "temp_screen.jpg"
    image.save(temp_image_path)
    try:
        model = YOLO("yolov5s")
        results = model.predict(source=temp_image_path)
        detected_objects = [result["name"] for result in results.pandas().xyxy[0].to_dict(orient="records")]
    except Exception as e:
        print(f"Error during object detection: {e}")
        detected_objects = ["No objects detected"]
    finally:
        if os.path.exists(temp_image_path):
            os.remove(temp_image_path)
    return detected_objects

def list_installed_programs():
    print("Installed applications:")
    software_list = get_installed_software()
    for software in software_list:
        print(software['name'])
    return [software['name'] for software in software_list]

def search_executables_in_paths(app_name, paths=SEARCH_PATHS):
    app_name_lower = app_name.lower()
    for path in paths:
        if path and os.path.exists(path):
            for root, _, files in os.walk(path):
                for file in files:
                    if app_name_lower in file.lower() and file.lower().endswith(".exe"):
                        return os.path.join(root, file)
    return None

def open_application(app_name):
    app_name = app_name.lower().strip()
    try:
        # Check installed software
        software_list = get_installed_software()
        software_names = [software['name'].lower() for software in software_list]

        # Exact match search
        for software in software_list:
            if app_name in software['name'].lower():
                install_location = software.get('install_location', None)
                if install_location:
                    executable_path = search_executables_in_paths("", [install_location])
                    if executable_path:
                        subprocess.Popen(executable_path, shell=True)
                        return f"Opening {software['name']}..."

        # Fallback: Search system directories for matching executables
        executable_path = search_executables_in_paths(app_name)
        if executable_path:
            subprocess.Popen(executable_path, shell=True)
            return f"Opening {app_name.capitalize()}..."

        # Look for a close match in the installed applications
        close_match = get_close_matches(app_name, software_names, n=1, cutoff=0.3)
        if close_match:
            matched_software = close_match[0]
            for software in software_list:
                if software['name'].lower() == matched_software:
                    install_location = software.get('install_location', None)
                    if install_location:
                        executable_path = search_executables_in_paths("", [install_location])
                        if executable_path:
                            subprocess.Popen(executable_path, shell=True)
                            return f"Opening {software['name']} (closest match to '{app_name}')..."

        return f"Could not find a program matching '{app_name}'."
    except Exception as e:
        print(f"Error opening program: {e}")
        return f"An error occurred while trying to open '{app_name}'."

def speak_text(text, rate=200):
    engine.setProperty("rate", rate)
    engine.say(text)
    engine.runAndWait()

def interpret_screen(user_command, conversation_history):
    print("Capturing screen...")
    screen_image = capture_screen()
    temp_image_path = "temp_screen.jpg"
    screen_image.save(temp_image_path)

    print("Extracting text from the screen...")
    screen_text = extract_text_from_image(screen_image)

    print("Detecting objects on the screen...")
    detected_objects = detect_objects(screen_image)

    if screen_text or detected_objects:
        screen_description = "Here's what I see based on my analysis:\n"
        if screen_text:
            screen_description += f"- Text visible on the screen: \"{screen_text}\".\n"
        if detected_objects:
            screen_description += f"- Detected items: {', '.join(detected_objects)}.\n"
    else:
        screen_description = "I couldn't detect any significant items on your screen."

    prompt = (
        f"Based on the following data, respond confidently as though you can see the screen.\n"
        f"The user asked: \"{user_command}\".\n"
        f"{screen_description}"
    )

    print(f"Prompt sent to GPT:\n{prompt}")

    conversation_history.append({"role": "user", "content": user_command})
    conversation_history.append({"role": "assistant", "content": screen_description})

    gpt_response = get_gpt_response(conversation_history)
    if gpt_response:
        print(f"GPT-4 Response: {gpt_response}")
        conversation_history.append({"role": "assistant", "content": gpt_response})
        speak_text(gpt_response)
    else:
        print("Failed to get a response from GPT.")

def get_gpt_response(conversation_history):
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=conversation_history
        )
        return response['choices'][0]['message']['content']
    except Exception as e:
        print(f"Error communicating with GPT: {e}")
        return None

def main():
    print("Welcome! Speak into your microphone, and I will respond.")
    print("Say 'exit' to quit, 'list applications' to see installed apps, or ask me to open an app.")

    while True:
        user_input = transcribe_speech(timeout=10)
        if user_input:
            if user_input.lower() == "exit":
                print("Goodbye!")
                break
            elif "list applications" in user_input.lower():
                installed_apps = list_installed_programs()
                speak_text("Here is the list of installed applications.")
            elif "open" in user_input.lower():
                app_name = user_input.lower().replace("open", "").strip()
                result = open_application(app_name)
                print(result)
                speak_text(result)
            elif "screen" in user_input.lower():
                interpret_screen(user_input, conversation_history)
            else:
                conversation_history.append({"role": "user", "content": user_input})
                response = get_gpt_response(conversation_history)
                if response:
                    print(f"GPT-4: {response}")
                    conversation_history.append({"role": "assistant", "content": response})
                    speak_text(response)
                else:
                    print("Failed to get a response from GPT.")
        else:
            print("No input detected. Try again.")

if __name__ == "__main__":
    main()
