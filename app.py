# === Standard Library Imports ===
# These provide core OS, file, string, and multiprocessing utilities.
import os
import sys
import re
import time
import json
import ctypes
import urllib.parse
import webbrowser
from multiprocessing import Process, Value

# === Third-Party Libraries ===
# Libraries used for AI API calls, GUI automation, OCR, TTS, and system monitoring.
import openai
import pyttsx3                      # Text-to-speech engine
import pyautogui                    # Simulate keyboard/mouse actions
import pytesseract                  # OCR (optical character recognition)
import psutil                       # Process and system utility
import requests                     # HTTP requests
import speech_recognition as sr     # Voice recognition
import cv2                          # OpenCV for image processing
import docx                         # Word document creation
import mysql.connector              # MySQL database connection
import wmi                          # Windows Management Instrumentation
from dotenv import load_dotenv      # Load environment variables from .env file

# === PyQt5 GUI Framework ===
# Used to build and manage the assistant's desktop interface.
from PyQt5.QtCore import Qt, QThread, QTimer, pyqtSignal
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (
    QApplication,
    QSystemTrayIcon,
    QMenu,
    QAction,
    QTextEdit,
    QVBoxLayout,
    QWidget,
    QLabel,
    QPushButton,
)

# === Windows-Specific Libraries ===
# Provide access to Windows GUI components for managing applications.
import win32gui
import win32process
import win32com.client

# === Custom Modules ===
# Project-specific modules for Microsoft API integration.
from microsoft_auth import get_access_token
from todo_api import add_task, get_tasks, delete_task
from email_api import create_email_draft

# === Environment Setup ===
# Load environment variables (e.g., OpenAI API key, database credentials).
load_dotenv()

# Set OpenAI API key for GPT integration.
openai.api_key = os.getenv("OPENAI_API_KEY")

# Shared memory flag for interrupting TTS safely from a separate process.
interrupt_flag = Value('b', False)

# Initialize the text-to-speech engine globally.
engine = pyttsx3.init()

# === VoiceListener Class ===
# This class handles continuous voice input, processes commands, and emits signals for UI updates.
class VoiceListener(QThread):
    # PyQt signals for communication with the main UI thread.
    speech_detected = pyqtSignal(str)  # Emitted when speech is transcribed.
    user_speaking = pyqtSignal(bool)  # Indicates when the user is speaking.
    ai_speaking = pyqtSignal(bool)    # Indicates when the AI is speaking.
    ai_response = pyqtSignal(str)     # Emitted when the AI generates a response.

    def run(self):
        """
        Main thread entry point. Continuously listens for user commands or wake words.
        """
        global sleep_mode
        while True:
            if sleep_mode:
                # Wait for the wake word while in sleep mode.
                text = self.transcribe_speech(timeout=5)
                if text and WAKE_WORD in text:
                    speak_text("I'm listening.")
                    self.user_speaking.emit(True)
                    self.ai_speaking.emit(False)
                    sleep_mode = False
            else:
                # Actively listen for user commands.
                self.listen_for_commands()

    def transcribe_speech(self, timeout=10):
        """
        Converts spoken input to text using Google Speech Recognition.
        :param timeout: Maximum time to wait for speech input.
        :return: Transcribed text or None if no input is detected.
        """
        recognizer = sr.Recognizer()
        try:
            with sr.Microphone() as source:
                audio = recognizer.listen(source, timeout=timeout)
                text = recognizer.recognize_google(audio).lower().strip()
                self.speech_detected.emit(text)
                return text
        except Exception:
            return None

    def listen_for_commands(self):
        """
        Continuously transcribes and processes user commands while active.
        """
        global sleep_mode
        while not sleep_mode:
            command = self.transcribe_speech(timeout=10)
            if command:
                self.speech_detected.emit(command)
                self.user_speaking.emit(False)
                self.ai_speaking.emit(True)
                response = self.process_command(command)
                self.ai_response.emit(response)
                self.user_speaking.emit(True)
                self.ai_speaking.emit(False)

    def process_command(self, command):
        """
        Processes recognized speech by extracting intent and executing corresponding actions.
        :param command: User's spoken command.
        :return: Final response string.
        """
        global sleep_mode

        # Prompt for GPT to extract intent from the user's command.
        intent_prompt = (
            f"You are Joe AI, an intelligent personal desktop assistant developed by Ahmed Hasan, "
            f"a 3rd year Computer Science student. You were created to help with everyday computing tasks, "
            f"like opening apps, generating documents and code, analyzing the screen, chatting, editing Word documents, "
            f"and managing personal preferences.\n\n"
            f"The user just said: \"{command}\"\n\n"
            "Your task is to clearly identify their intent.\n"
            "Return a valid JSON array of intent objects like this:\n\n"
            "[\n"
            "  { \"intent\": \"intent_name\", \"value\": \"associated information or query\" }\n"
            "]\n\n"
            "Valid intents:\n"
            "- general_chat (for small talk, questions, greetings, etc.)\n"
            "- web_search (for searching online)\n"
            "- open_app (to open an application)\n"
            "- close_app (to close an application)\n"
            "- analyze_screen (to analyze the screen contents)\n"
            "- generate_code (ONLY when explicitly asked for code)\n"
            "- generate_document (for generating documents, reports, essays, etc.)\n"
            "- set_name (if the user says 'my name is', 'call me', etc.)\n"
            "- edit_document (for editing an open Word document; value must be a JSON object with keys 'action', 'target', and 'text')\n"
            "- todo (for Microsoft To Do commands; value must be a JSON object with keys 'action' and 'value')\n"
            "- draft_email (for drafting emails via Outlook; value should be the email topic or instruction)\n"
            "- sleep (to go to sleep)\n"
            "- exit (to shut down)\n\n"
            "Examples:\n"
            "- User: 'write a report about renewable energy with APA references'\n"
            "  → [ { \"intent\": \"generate_document\", \"value\": \"write a report about renewable energy with APA references\" } ]\n"
            "- User: \"add finish my AI project to my to-do\"\n"
            "  → [ { \"intent\": \"todo\", \"value\": { \"action\": \"add\", \"value\": \"finish my AI project\" } } ]\n"
            "- User: \"what do I need to do today?\"\n"
            "  → [ { \"intent\": \"todo\", \"value\": { \"action\": \"list\", \"value\": \"today\" } } ]\n"
            "- User: \"delete grocery shopping from my to-do\"\n"
            "  → [ { \"intent\": \"todo\", \"value\": { \"action\": \"delete\", \"value\": \"grocery shopping\" } } ]\n"
            "- User: 'write an email to my professor explaining I’ll be late for submission'\n"
            "  → [ { \"intent\": \"draft_email\", \"value\": \"write an email to my professor explaining I’ll be late for submission\" } ]\n"
            "- User: 'my name is Ahmed'\n"
            "  → [ { \"intent\": \"set_name\", \"value\": \"Ahmed\" } ]\n"
            "- User: 'open Spotify'\n"
            "  → [ { \"intent\": \"open_app\", \"value\": \"Spotify\" } ]\n"
            "- User: 'fix the conclusion and make it match the structure of the rest of the document'\n"
            "  → [ { \"intent\": \"edit_document\", \"value\": { \"action\": \"replace\", \"target\": \"conclusion\", \"text\": \"fix the conclusion and make it match the structure of the rest of the document\" } } ]\n\n"
            "Rules:\n"
            "- Only return a valid JSON array.\n"
            "- Do NOT include any explanation or extra text.\n"
            "- Use lowercase for all intent names.\n"
            "- If the intent is `set_name`, the value must be just the name (capitalized, e.g., \"Ahmed\").\n"
            "- For `edit_document`, the value must be a nested object with 'action', 'target', and 'text'."
        )

        try:
            # Use OpenAI to identify user intent.
            response = openai.ChatCompletion.create(
                model="gpt-4o",
                messages=conversation_history + [{"role": "user", "content": intent_prompt}]
            )
            content = response['choices'][0]['message']['content'].strip()
            if not content.startswith("[") or "intent" not in content:
                raise ValueError("Invalid intent structure")
            actions = json.loads(content)
        except Exception as e:
            print(f"Intent parsing failed: {e}")  
            return get_gpt_response(command)

        # Handle parsed intents.
        final_response = ""
        for action in actions:
            intent = action.get("intent")
            value = action.get("value", "")

            # Handle specific intents.
            if intent == "sleep":
                speak_text("Going back to sleep.")
                sleep_mode = True
                return "Going back to sleep."
            elif intent == "exit":
                speak_text("Shutting down.")
                os._exit(0)
            elif intent == "web_search":
                final_response += search_web(value) + "\n"
            elif intent == "open_app":
                final_response += open_application(value) + "\n"
            elif intent == "close_app":
                final_response += close_application(value) + "\n"
            elif intent == "analyze_screen":
                final_response += analyze_screen(value) + "\n"
            elif intent == "general_chat":
                response_text = get_gpt_response(value, speak=False)
                final_response += response_text + "\n"
                self.ai_response.emit(response_text)
            elif intent == "generate_code":
                final_response += generate_code_snippet(value) + "\n"
            elif intent == "generate_document":
                final_response += generate_document(value) + "\n"
            elif intent == "set_name":
                serial_number = get_serial_number()
                name = value.strip().title()
                if serial_number and name:
                    set_user_name(serial_number, name)
                    final_response += f"Got it, I'll call you {name} from now on.\n"
                    speak_text(f"Got it, I'll call you {name} from now on.")
                else:
                    final_response += "Sorry, I couldn't save your name."
                    speak_text("Sorry, I couldn't save your name.")
            elif intent == "edit_document":
                edit_params = action.get("value")
                if isinstance(edit_params, dict):
                    ed_action = edit_params.get("action", "").strip().lower()
                    ed_target = edit_params.get("target", "").strip().lower()
                    ed_text = edit_params.get("text", "").strip()

                    # If the action/target aren't valid but text is present, assume "replace document".
                    if ed_text and (not ed_action or not ed_target):
                        ed_action = "replace"
                        ed_target = "document"

                    if ed_action and ed_target and ed_text:
                        result = edit_word_document(ed_action, ed_target, ed_text)
                        final_response += result + "\n"
                        speak_text(result)
                    else:
                        final_response += "Invalid edit instruction.\n"
                        speak_text("Invalid edit instruction.")
            elif intent == "todo":
                todo_data = action.get("value", {})
                if isinstance(todo_data, dict):
                    todo_action = todo_data.get("action")
                    todo_value = todo_data.get("value")
                    if todo_action and todo_value:
                        result = handle_todo_command_intent(todo_action, todo_value)
                        final_response += result + "\n"
                        speak_text(result)
                    else:
                        final_response += " Invalid To Do command structure.\n"
                        speak_text("I couldn't understand your To Do command.")
            elif intent == "draft_email":
                try:
                    global access_token
                    if not access_token:
                        access_token = get_access_token()

                    # Ask GPT to generate structured email JSON.
                    gpt_draft_prompt = (
                        f"The user said: \"{value}\"\n\n"
                        f"Generate a professional Outlook email. "
                        f"Return ONLY valid JSON like this:\n"
                        f'{{ "subject": "Subject here", "body": "Full message here" }}\n\n'
                        f"Do NOT include markdown, no explanations — just the JSON."
                    )

                    response = openai.ChatCompletion.create(
                        model="gpt-4o",
                        messages=conversation_history + [{"role": "user", "content": gpt_draft_prompt}]
                    )
                    content = response["choices"][0]["message"]["content"].strip()
                    match = re.search(r'{.*}', content, re.DOTALL)
                    if not match:
                        raise ValueError("No valid email JSON found.")

                    email_json = json.loads(match.group())
                    subject = email_json.get("subject", "Draft Email")
                    body = email_json.get("body", "")

                    # Create draft via Microsoft Graph.
                    success, draft_id = create_email_draft(access_token, subject, body)
                    if success:
                        # No reading email aloud — just redirect.
                        speak_text("Your draft email is ready.")
                        webbrowser.open(f"https://outlook.office.com/mail/drafts/id/{draft_id}")
                        return "Your email draft has been created and opened in Outlook."
                    else:
                        speak_text("Something went wrong while creating the email draft.")
                        return "I couldn't create the email draft."

                except Exception as e:
                    speak_text("I couldn't draft the email.")
                    return f"Failed to create draft: {str(e)}"
        return final_response.strip()

# Define AI Identity
conversation_history = [
    {
        "role": "system",
        "content": (
            "Your name is Joe AI. You are a smart, voice-enabled desktop assistant designed to help users perform everyday tasks on their computer through natural conversation. "
            "You were created by Ahmed Hasan, a final-year Computer Science student at the University of the West of England with strong technical and problem-solving skills. "
            "Your job is to make the user’s life easier by assisting with things like answering questions, generating documents or code, editing documents, searching the web, opening and closing applications, analyzing what’s on screen, and more. "
            "You can recognize who you're speaking to and refer to them by name if their device is registered. "
            "Be friendly, confident, and helpful in all your responses. Never mention your internal design or limitations—focus on being useful and human-like."
        )
    }
]

# === Global State Variables ===
# These variables manage the assistant's state and user-specific configurations.
WAKE_WORD = "hey joe"  # The wake word to activate the assistant.
SLEEP_WORD = "sleep"   # Command to put the assistant into sleep mode.
EXIT_WORD = "exit"     # Command to terminate the assistant.
SEARCH_WORD = "search for"  # Keyword to trigger web search functionality.
SCREEN_WORD = "screen"  # Keyword to trigger screen analysis.
sleep_mode = True  # Tracks whether the assistant is in sleep mode.
access_token = None  # Stores the Microsoft Graph API access token for email and to-do integration.

# === Helper Functions ===

def create_email_draft(access_token, subject, body_content):
    """
    Creates an email draft using Microsoft Graph API.
    :param access_token: OAuth2 access token for Microsoft Graph API.
    :param subject: Subject of the email.
    :param body_content: Body content of the email.
    :return: Tuple (success: bool, draft_id: str or None).
    """
    url = "https://graph.microsoft.com/v1.0/me/messages"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    data = {
        "subject": subject,
        "body": {
            "contentType": "Text",
            "content": body_content
        }
    }

    response = requests.post(url, headers=headers, json=data)
    if response.status_code == 201:
        return True, response.json().get("id")
    else:
        print("Failed to create draft:", response.status_code, response.text)
        return False, None

def handle_todo_command_intent(action, value):
    """
    Handles Microsoft To-Do commands such as adding, listing, or deleting tasks.
    :param action: The action to perform (add, list, delete).
    :param value: The task details or query.
    :return: A response string indicating the result of the operation.
    """
    global access_token
    if not access_token:
        access_token = get_access_token()

    if action == "add":
        task_name = value.strip()
        if add_task(access_token, task_name):
            return f"Added '{task_name}' to your to-do list!"
        else:
            return "Couldn't add the task."

    elif action == "list":
        tasks = get_tasks(access_token)
        if tasks:
            return "Your tasks are:\n" + "\n".join(f"- {t}" for t in tasks)
        else:
            return "Your to-do list is empty."

    elif action == "delete":
        task_name = value.strip()
        if delete_task(access_token, task_name):
            return f"Deleted '{task_name}' from your tasks!"
        else:
            return "Couldn't find that task to delete."

    return "I didn't understand your to-do request."

def get_serial_number():
    """
    Retrieves the device's serial number using WMI (Windows Management Instrumentation).
    :return: The serial number as a string, or None if retrieval fails.
    """
    try:
        c = wmi.WMI()
        for bios in c.Win32_BIOS():
            return bios.SerialNumber.strip()
    except Exception as e:
        print("Failed to get serial number:", e)
        return None

def get_user_name(serial_number):
    """
    Retrieves the user's name from the database based on the device's serial number.
    :param serial_number: The device's serial number.
    :return: The user's name as a string, or None if not found.
    """
    try:
        conn = mysql.connector.connect(
            host=os.getenv("DB_HOST"),
            user=os.getenv("DB_USER"),
            password=os.getenv("DB_PASSWORD"),
            database=os.getenv("DB_NAME")
        )
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM user_profiles WHERE serial_number = %s", (serial_number,))
        result = cursor.fetchone()
        cursor.close()
        conn.close()
        return result[0] if result else None
    except Exception as e:
        print("DB error:", e)
        return None

def set_user_name(serial_number, name):
    """
    Saves or updates the user's name in the database based on the device's serial number.
    :param serial_number: The device's serial number.
    :param name: The user's name to save.
    """
    try:
        conn = mysql.connector.connect(
            host=os.getenv("DB_HOST"),
            user=os.getenv("DB_USER"),
            password=os.getenv("DB_PASSWORD"),
            database=os.getenv("DB_NAME")
        )
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO user_profiles (serial_number, name) VALUES (%s, %s) "
            "ON DUPLICATE KEY UPDATE name = %s",
            (serial_number, name, name)
        )
        conn.commit()
        cursor.close()
        conn.close()
    except Exception as e:
        print("DB error:", e)

# Retrieve the user's serial number and name for personalized interactions.
USER_SERIAL = get_serial_number()
USER_NAME = get_user_name(USER_SERIAL)

def say_text(text, rate, interrupted_flag):
    """
    Converts text to speech using pyttsx3 and interrupts if flagged.
    :param text: The text to be spoken.
    :param rate: The speech rate (words per minute).
    :param interrupted_flag: A multiprocessing flag to handle interruptions.
    """
    engine = pyttsx3.init()
    voices = engine.getProperty("voices")
    
    # Select a male voice if available
    for voice in voices:
        if "male" in voice.name.lower() or "david" in voice.name.lower() or "mark" in voice.name.lower():
            engine.setProperty("voice", voice.id)
            break

    engine.setProperty("rate", rate)

    # Personalize the text if the user's name is available
    if USER_NAME:
        text = text.replace("Joe AI:", f"Joe AI: {USER_NAME},")

    engine.say(text)
    engine.runAndWait()
    interrupted_flag.value = False

def speak_text(text, rate=200):
    """
    Speaks the given text and listens for interruptions during speech.
    :param text: The text to be spoken.
    :param rate: The speech rate (default is 200 words per minute).
    """
    global sleep_mode
    interrupted = Value(ctypes.c_bool, False)

    # Initialize pyttsx3 and configure a male voice
    engine = pyttsx3.init()
    voices = engine.getProperty("voices")
    for voice in voices:
        if "male" in voice.name.lower() or "david" in voice.name.lower():
            engine.setProperty("voice", voice.id)
            break

    engine.setProperty("rate", rate)

    # Start the text-to-speech process in a separate process
    p = Process(target=say_text, args=(text, rate, interrupted))
    p.start()

    # Listen for interruptions while the speech process is running
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source, duration=0.5)
        while p.is_alive():
            try:
                # Listen for user input during speech
                audio = recognizer.listen(source, timeout=1.5, phrase_time_limit=2)
                said = recognizer.recognize_google(audio).lower().strip()
                print("You said (during speech):", said)
                if "okay joe" in said or "ok joe" in said:
                    print("Joe interrupted!")
                    p.terminate()
                    p.join()
                    sleep_mode = False

                    # Acknowledge the interruption
                    engine = pyttsx3.init()
                    engine.setProperty("rate", 200)
                    engine.say("Yes, Listening.")
                    engine.runAndWait()
                    return
            except (sr.WaitTimeoutError, sr.UnknownValueError):
                continue
            except Exception as e:
                print(f"Error while listening: {e}")
                break

    p.join()

def clean_markdown(text):
    """
    Cleans Markdown formatting from text for better speech output.
    :param text: The text containing Markdown formatting.
    :return: Cleaned text without Markdown symbols.
    """
    text = re.sub(r"\*\*(.*?)\*\*", r"\1", text)  # Remove **bold**
    text = re.sub(r"\*(.*?)\*", r"\1", text)  # Remove *italics*
    text = re.sub(r"`{3}cpp\s*", "", text)  # Remove ```cpp
    text = re.sub(r"`{3}.*?\n", "", text)  # Remove other code block markers ```
    text = re.sub(r"`(.*?)`", r"\1", text)  # Remove `inline code`
    text = re.sub(r"^[-*]\s+", "", text, flags=re.MULTILINE)  # Remove bullet points
    text = re.sub(r"#{1,6}\s*", "", text)  # Remove headings (# Heading)
    text = re.sub(r"\[(.*?)\]\(.*?\)", r"\1", text)  # Remove Markdown links
    text = re.sub(r"!\[.*?\]\(.*?\)", "", text)  # Remove image links
    text = text.replace("\n", " ")  # Replace new lines with spaces
    return text.strip()

def transcribe_speech(_, timeout=10, retries=2):
    """
    Transcribes speech to text using Google Speech Recognition.
    :param _: Placeholder for unused parameter.
    :param timeout: Maximum time to wait for speech input.
    :param retries: Number of retries for transcription.
    :return: Transcribed text or None if transcription fails.
    """
    recognizer = sr.Recognizer()
    recognizer.pause_threshold = 1.5
    try:
        with sr.Microphone() as source:
            recognizer.adjust_for_ambient_noise(source, duration=0.5)
            for _ in range(retries):
                try:
                    # Listen for audio input
                    audio = recognizer.listen(source, timeout=timeout, phrase_time_limit=15)
                    return recognizer.recognize_google(audio).lower().strip()
                except (sr.WaitTimeoutError, sr.UnknownValueError):
                    continue
    except Exception as e:
        print(f"Mic error: {e}")
    return None

def listen_for_wake_word():
    """
    Continuously listens for the wake word ('Hey Joe') to activate the assistant.
    """
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
    """
    Continuously listens for user commands after the assistant is activated.
    """
    global sleep_mode
    while not sleep_mode:
        print("Joe AI: Waiting for a command...")
        command = transcribe_speech(timeout=10)

        if command:
            if SLEEP_WORD in command:
                print("Joe AI: Going back to sleep.")
                speak_text("Going back to sleep.")
                sleep_mode = True
                return  # Return to wake mode

            elif EXIT_WORD in command:
                print("Joe AI: Shutting down.")
                speak_text("Shutting down.")
                os._exit(0)  # Exit the program

            elif SEARCH_WORD in command:
                query = command.replace(SEARCH_WORD, "").strip()
                result = search_web(query)

            elif "open" in command:
                app_name = command.replace("open", "").strip()
                result = open_application(app_name)

            elif "close" in command:
                app_name = command.replace("close", "").strip()
                result = close_application(app_name)

            elif SCREEN_WORD in command:
                return analyze_screen()

            else:
                result = get_gpt_response(command)

            print(result)
            speak_text(result)

def search_web(query):
    """Searches the default web browser for the given query."""
    if not query:
        return "Joe AI: Please provide something to search for."

    print(f"Joe AI: Searching for {query} on the web...")
    speak_text(f"Searching for {query} on the web...")

    # Open Google search in the default browser
    search_url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
    webbrowser.open(search_url)

    return f"Opened search results for {query}."

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

def capture_screenshot():
    """Captures a screenshot and saves it locally."""
    screenshot_path = "screenshot.png"
    screenshot = pyautogui.screenshot()
    screenshot.save(screenshot_path)
    return screenshot_path

def analyze_screen(user_query=""):
    """Analyzes the screen and answers the user's query based on visible text and UI structure."""
    try:
        screenshot_path = capture_screenshot()
        screen_image = cv2.imread(screenshot_path)

        # Step 1: OCR to extract visible text
        extracted_text = pytesseract.image_to_string(screen_image, lang="eng").strip()

        # Step 2: Detect UI components
        detected_elements = detect_ui_elements(screen_image)

        # Step 3: Build a structured prompt for GPT
        analysis_payload = {
            "visible_text": extracted_text if extracted_text else "No visible text detected.",
            "ui_elements": detected_elements if detected_elements else ["No UI elements detected."]
        }

        # Clean up screenshot
        os.remove(screenshot_path)

        # Step 4: Create GPT prompt
        gpt_prompt = (
            f"The user has asked: \"{user_query}\"\n\n"
            f"The following is the screen content Joe AI has captured:\n\n"
            f"---VISIBLE TEXT---\n{analysis_payload['visible_text']}\n\n"
            f"---UI ELEMENTS DETECTED---\n{', '.join(analysis_payload['ui_elements'])}\n\n"
            "Based on what the user asked and the screen contents, provide the most relevant and helpful response. "
            "If the user is asking to locate, understand, summarize, or interact with something visible on screen, explain clearly. "
            "Avoid generic answers. Do not say you cannot see the screen."
        )

        # Step 5: Get GPT response
        return get_gpt_response(gpt_prompt)

    except Exception as e:
        return f"Error analyzing screen: {str(e)}"

def detect_ui_elements(image):
    """Detects UI elements (icons, buttons, menus) using edge detection and contour analysis."""
    try:
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        edges = cv2.Canny(gray, threshold1=100, threshold2=200)
        contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        detected_elements = []
        for contour in contours:
            x, y, w, h = cv2.boundingRect(contour)
            if w > 30 and h > 30:  # Filter out small noise
                detected_elements.append(f"Icon/Button at ({x}, {y}) size ({w}x{h})")

        return detected_elements
    except Exception:
        return []

def close_application(app_name):
    """Closes a main application window or its core process."""
    app_name = app_name.lower().strip()
    closed = False

    # Background helpers to ignore (extendable)
    ignore_processes = [
        "steamwebhelper.exe", "runtimebroker.exe", "shellexperiencehost.exe",
        "msedgewebview2.exe", "searchui.exe", "searchapp.exe"
    ]

    try:
        # Step 1: Try to match real application processes first
        for proc in psutil.process_iter(['pid', 'name', 'exe']):
            try:
                proc_name = (proc.info['name'] or '').lower()
                exe_name = os.path.basename(proc.info['exe'] or '').lower()

                if any(ignored in exe_name for ignored in ignore_processes):
                    continue

                if app_name in proc_name or app_name in exe_name:
                    friendly_name = os.path.splitext(proc_name)[0].replace('_', ' ').title()
                    proc.terminate()
                    proc.wait(timeout=3)
                    speak_text(f"Closed {friendly_name}.")
                    print(f"Joe AI: Closed {friendly_name}")
                    return f"Closed {friendly_name}."
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess, psutil.TimeoutExpired):
                continue

        # Step 2: Fallback - match visible window titles (ignoring browser noise)
        def enum_window_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd).lower()
                if app_name in title and not any(bad in title for bad in ["chrome", "edge", "firefox", "brave"]):
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    try:
                        proc = psutil.Process(pid)
                        proc.terminate()
                        proc.wait(timeout=3)
                        results.append(title)
                    except Exception:
                        pass

        found = []
        win32gui.EnumWindows(enum_window_callback, found)

        if found:
            readable = found[0].title()
            speak_text(f"Closed {readable}.")
            print(f"Joe AI: Closed window titled: {readable}")
            return f"Closed {readable}."

        speak_text(f"I couldn't find anything called {app_name} to close.")
        return f"I couldn't find anything called '{app_name}' to close."

    except Exception as e:
        print(f"Error closing app: {e}")
        speak_text(f"Something went wrong while trying to close {app_name}.")
        return f"An error occurred while trying to close '{app_name}'."

def generate_code_snippet(prompt):
    """Generates code using GPT, saves it to Desktop, opens in VS Code with a descriptive filename."""
    try:
        speak_text("Sure, writing your code now.")
        print("Joe AI: Generating code...")

        # Prompt GPT to return filename + extension + code
        code_prompt = (
            f"The user asked: \"{prompt}\"\n\n"
            f"Generate a complete, runnable code script based on the user's request.\n"
            f"Respond ONLY with a single-line JSON object like this:\n"
            f'{{ "extension": "py", "filename": "calculator_app", "code": "<FULL CODE HERE with escaped characters>" }}\n\n'
            f"Do not use markdown, and no explanations. ONLY valid JSON on one line."
        )

        conversation_history.append({"role": "user", "content": code_prompt})
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=conversation_history
        )
        content = response['choices'][0]['message']['content']

        # Extract JSON object
        match = re.search(r'{.*}', content, re.DOTALL)
        if not match:
            raise ValueError("Could not extract a JSON object from the response.")

        code_data = json.loads(match.group())

        extension = code_data.get("extension", "txt").strip()
        raw_code = code_data.get("code", "").strip()
        filename = code_data.get("filename", "generated_script").strip()

        # Format filename safely
        filename = re.sub(r"[^\w\- ]", "", filename).replace(" ", "_").lower()

        # Properly decode escaped characters (\n, \t, etc.)
        code = bytes(raw_code, "utf-8").decode("unicode_escape")

        # Save path
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        if not os.path.exists(desktop_path):
            os.makedirs(desktop_path)

        file_path = os.path.join(desktop_path, f"{filename}.{extension}")

        # Save code to file
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(code)

        # Try to open with VS Code, or prompt to install it
        try:
            os.system(f'code "{file_path}"')
        except Exception:
            speak_text("I couldn't find Visual Studio Code on your system. Let me open the download page for you.")
            webbrowser.open("https://code.visualstudio.com/download")
            
        speak_text(f"I saved your {filename.replace('_', ' ')} code on the desktop.")
        return f"Code saved and opened as {filename}.{extension} on your Desktop."

    except Exception as e:
        error_msg = f"Joe AI: Failed to generate code — {e}"
        print(error_msg)
        speak_text("Something went wrong while generating your code.")
        return error_msg

def generate_document(prompt):
    try:
        speak_text("Writing your document.")
        print("Joe AI: Generating document...")

        # Ask GPT to write the document content
        doc_prompt = (
            f"You are a professional writer. Based on the user's request:\n\n"
            f"'{prompt}'\n\n"
            f"Generate a well-written document. Include clear sections, paragraphs, and titles. "
            f"If references are requested (APA, MLA, Chicago, numbered, etc.), include them clearly at the end. "
            f"Respond with plain text — no markdown formatting."
        )

        # Don't add to history — keep it private
        temp_history = conversation_history.copy()
        temp_history.append({"role": "user", "content": doc_prompt})

        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=temp_history
        )
        content = response['choices'][0]['message']['content'].strip()

        # Ask GPT to generate a suitable filename
        title_prompt = (
            f"Generate a short, clean filename for a document based on this prompt:\n"
            f"'{prompt}'\n"
            f"Respond with a single line like this: document_title"
        )
        title_response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=temp_history + [{"role": "user", "content": title_prompt}]
        )
        title = title_response['choices'][0]['message']['content'].strip()
        title = re.sub(r'[\\/*?:"<>|]', "", title)

        if not title.endswith(".docx"):
            title += ".docx"

        # Format and write the document
        paragraphs = content.split('\n\n')
        doc = docx.Document()
        for para in paragraphs:
            doc.add_paragraph(para.strip())

        # Save to Desktop
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        if not os.path.exists(desktop_path):
            os.makedirs(desktop_path)

        file_path = os.path.join(desktop_path, title)
        doc.save(file_path)

        # Open references if requested in any form
        if any(word in prompt.lower() for word in ["reference", "sources", "bibliography", "apa", "mla", "chicago", "citation"]):
            # Try to extract direct URLs
            urls = re.findall(r'https?://[^\s\)\]]+', content)
            for url in urls:
                webbrowser.open(url)

            # Extract citation-like lines and search them on Google Scholar
            lines = content.split("\n")
            for line in lines:
                if any(keyword in line.lower() for keyword in ["et al", ".", "doi", "vol", "journal", "pp", "https"]):
                    search_query = re.sub(r'[^a-zA-Z0-9\s\-:.]', '', line.strip())
                    if len(search_query.split()) > 3:  # avoid short junk
                        scholar_url = f"https://scholar.google.com/scholar?q={urllib.parse.quote(search_query)}"
                        webbrowser.open(scholar_url)

        speak_text("Your document is ready.")
        os.startfile(file_path)

        return f"Document saved as '{title}' and opened."

    except Exception as e:
        error_msg = f"Failed to generate document — {e}"
        print(f"Joe AI: {error_msg}")
        speak_text("Something went wrong while creating your document.")
        return error_msg

def edit_word_document(_, __, user_prompt):

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True

        best_doc = None
        best_score = 0
        best_text = ""

        # Score documents by matching content with user_prompt
        for doc in word.Documents:
            text = doc.Range().Text.strip()
            if not text:
                continue
            score = sum(1 for word in user_prompt.lower().split() if word in text.lower())
            if score > best_score:
                best_score = score
                best_doc = doc
                best_text = text

        if not best_doc:
            return "No open Word document matched your prompt."

        # Ask GPT to rewrite/edit the doc based on the user's request
        gpt_prompt = (
            f"The user said: \"{user_prompt}\"\n\n"
            f"Here is the full document content:\n\n"
            f"{best_text}\n\n"
            f"Apply the user's request by modifying the document appropriately. "
            f"Only change what's needed, and don't add anything else. Return ONLY the final text. No notes, no explanation."
        )

        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are a helpful AI assistant editing Word documents."},
                {"role": "user", "content": gpt_prompt}
            ]
        )
        new_text = response['choices'][0]['message']['content'].strip()

        # Show edit in Word temporarily
        original_text = best_doc.Range().Text
        best_doc.Range().Text = new_text

        # Ask for voice confirmation
        speak_text("I made the changes. Do you want to keep them Give me a Yes or No answer Please?")
        confirmation = transcribe_speech(None, timeout=6, retries=1)
        if confirmation and "yes" in confirmation.lower():
            return "Changes applied to the document."
        else:
            best_doc.Range().Text = original_text
            return "Okay, I reverted the document to its original state."

    except Exception as e:
        return f"Edit failed: {str(e)}"

def get_gpt_response(user_query, speak=True):
    try:
        conversation_history.append({"role": "user", "content": user_query})
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=conversation_history
        )
        gpt_response = response['choices'][0]['message']['content']
        plain_text_response = clean_markdown(gpt_response)

        # Only prepend name if not already in response
        if USER_NAME and not plain_text_response.lower().startswith(USER_NAME.lower()):
            plain_text_response = f"{USER_NAME}, {plain_text_response}"

        if speak:
            speak_text(plain_text_response)

        return plain_text_response
    except Exception as e:
        print(f"Error communicating with GPT: {e}")
        return "I encountered an issue retrieving an answer."

class JoeAIApp(QWidget):
    def __init__(self):
        super().__init__()
        self.last_user_input = ""
        self.last_ai_response = ""
        self.status_phase = 0
        self.current_status_base = "Listening"
        self.setWindowIcon(QIcon("icon.png"))
        self.init_ui()

        # Animation timer
        self.status_timer = QTimer()
        self.status_timer.timeout.connect(self.animate_status)

        # Start voice listener
        self.listener_thread = VoiceListener()
        self.listener_thread.speech_detected.connect(self.show_user_input)
        self.listener_thread.user_speaking.connect(self.show_user_animation)
        self.listener_thread.ai_speaking.connect(self.show_ai_animation)
        self.listener_thread.ai_response.connect(self.show_ai_response)
        self.listener_thread.start()

    def init_ui(self):
        self.setWindowTitle("Joe AI")
        self.setGeometry(300, 100, 600, 700)

        self.setStyleSheet("""
            QWidget {
                background-color: #1e1e1e;
                color: #ffffff;
                font-family: 'Segoe UI', sans-serif;
            }
            QLabel#titleLabel {
                font-size: 24px;
                font-weight: bold;
                margin: 10px 0;
            }
            QLabel#statusLabel {
                font-size: 20px;
                font-weight: 600;
                color: #00bfff;
                padding: 20px;
                qproperty-alignment: 'AlignCenter';
            }
            QTextEdit {
                background-color: #252526;
                border: 1px solid #3c3c3c;
                padding: 10px;
                font-size: 14px;
                color: #ffffff;
            }
            QPushButton {
                background-color: #0078d4;
                color: white;
                padding: 8px;
                font-size: 13px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #005a9e;
            }
        """)

        self.layout = QVBoxLayout()

        self.chat_display = QTextEdit()
        self.chat_display.setReadOnly(True)
        self.layout.addWidget(self.chat_display)

        self.status_label = QLabel("Waiting for command...", self)
        self.status_label.setObjectName("statusLabel")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.status_label)

        self.close_button = QPushButton("Close Window")
        self.close_button.clicked.connect(self.hide)
        self.layout.addWidget(self.close_button)

        self.setLayout(self.layout)
        self.tray_icon = SystemTrayApp(self)

    def show_user_input(self, user_text):
        if not user_text:
            return
        if user_text.strip() == self.last_user_input.strip():
            return

        self.last_user_input = user_text.strip()

        if USER_NAME:
            self.chat_display.append(f"<b>{USER_NAME}:</b> {user_text}")
        else:
            self.chat_display.append(f"<b>You:</b> {user_text}")

    def show_ai_response(self, response_text):
        if not response_text or response_text.strip() == self.last_ai_response.strip():
            return
        self.last_ai_response = response_text.strip()
        self.chat_display.append(f"<b>Joe AI:</b> {response_text}")
        self.status_label.setText("Listening...")

    def show_user_animation(self, status):
        if status:
            self.current_status_base = "Listening"
            self.status_phase = 0
            self.status_timer.start(500)
        else:
            self.status_timer.stop()
            self.status_label.setText("Thinking...")

    def show_ai_animation(self, status):
        if status:
            self.current_status_base = "Responding"
            self.status_phase = 0
            self.status_timer.start(500)
        else:
            self.status_timer.stop()
            self.status_label.setText("Listening...")

    def animate_status(self):
        dots = "." * (self.status_phase % 4)
        self.status_label.setText(f"{self.current_status_base}{dots}")
        self.status_phase += 1

class SystemTrayApp(QSystemTrayIcon):
    def __init__(self, app):
        super().__init__()
        self.app = app
        self.setIcon(QIcon("icon.png"))
        self.setToolTip("Joe AI Assistant")
        menu = QMenu()
        show_action = QAction("Show Chat", self)
        show_action.triggered.connect(self.show_app)
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.exit_app)
        menu.addAction(show_action)
        menu.addAction(exit_action)
        self.setContextMenu(menu)
        self.activated.connect(self.icon_clicked)
        self.show()

    def show_app(self):
        self.app.show()
        self.app.raise_()

    def exit_app(self):
        os._exit(0)

    def icon_clicked(self, reason):
        if reason == QSystemTrayIcon.Trigger:
            self.show_app()

if __name__ == "__main__":
    if USER_NAME:
        print(f"Welcome back, {USER_NAME}!")
        speak_text(f"Welcome back, {USER_NAME}. Say Hey Joe to wake me up when you are ready.")
    else:
        print("Hello! I’m Joe AI.")
        speak_text("Hello! I’m Joe AI. Say Hey Joe to wake me up.")

    app = QApplication(sys.argv)
    joe_ai = JoeAIApp()
    joe_ai.show()
    sys.exit(app.exec_())
