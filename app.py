import os
import webbrowser
import openai
import pyttsx3
import pyautogui
import pytesseract
import numpy as np
import speech_recognition as sr
import psutil
import time
from PyQt5.QtWidgets import QApplication, QSystemTrayIcon, QMenu, QAction, QTextEdit, QVBoxLayout, QWidget, QLabel, QPushButton, QHBoxLayout
from PyQt5.QtGui import QIcon, QMovie
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import sys
import re
import cv2
import json
import docx
import urllib.parse
import mysql.connector
import wmi

# OpenAI API key
openai.api_key = "sk-tbtbdjXN0PNeKVX8x6oXJFABUkwYsEeOj9TinWn3jOT3BlbkFJuGto6skfATpazIFkDBnEr1JtKDe0ykgJkavseRQP0A"

# Initialize text-to-speech engine
engine = pyttsx3.init()
class VoiceListener(QThread):
    speech_detected = pyqtSignal(str)
    user_speaking = pyqtSignal(bool)
    ai_speaking = pyqtSignal(bool)
    log_message = pyqtSignal(str)
    ai_response = pyqtSignal(str)
    
    def run(self):
        global sleep_mode
        while True:
            if sleep_mode:
                text = self.transcribe_speech(timeout=5)
                if text and WAKE_WORD in text:
                    self.log_message.emit("I'm listening!")
                    speak_text("I'm listening.")
                    self.user_speaking.emit(True)
                    self.ai_speaking.emit(False)
                    sleep_mode = False
            else:
                self.listen_for_commands()
    
    def transcribe_speech(self, timeout=10):
        recognizer = sr.Recognizer()
        try:
            with sr.Microphone() as source:
                self.log_message.emit("Listening...")
                audio = recognizer.listen(source, timeout=timeout)
                text = recognizer.recognize_google(audio).lower().strip()
                self.log_message.emit(f"You said: {text}")
                return text
        except Exception as e:
            self.log_message.emit(f"Error: {e}")
            return None
    
    def listen_for_commands(self):
        global sleep_mode
        while not sleep_mode:
            command = self.transcribe_speech(timeout=10)
            if command:
                self.log_message.emit(f"Processing command: {command}")
                self.speech_detected.emit(command)
                self.user_speaking.emit(False)
                self.ai_speaking.emit(True)
                response = self.process_command(command)
                self.ai_response.emit(response)
                self.log_message.emit(f"Joe AI: {response}")
                self.user_speaking.emit(True)
                self.ai_speaking.emit(False)

    def process_command(self, command):
        global sleep_mode

        intent_prompt = (
            f"You are Joe AI. The user said: \"{command}\".\n\n"
            "Identify the user's intent clearly. Respond with ONLY a JSON array of objects. Each object should follow this format:\n"
            "[\n"
            "  { \"intent\": \"generate_document\", \"value\": \"write a report about the sun with harvard references\" }\n"
            "]\n\n"
            "Valid intents:\n"
            "- general_chat (for greetings, thanks, questions, or small talk)\n"
            "- web_search\n"
            "- open_app\n"
            "- close_app\n"
            "- analyze_screen\n"
            "- generate_code (only if the user asks to write, create, or generate code)\n"
            "- generate_document (for reports, articles, essays, etc.)\n"
            "- set_name (if the user says things like 'my name is [name]' or 'call me [name]')\n"
            "- sleep\n"
            "- exit\n\n"
            "Use `generate_document` if the user says write/generate/create a report, article, essay, document, or mentions references or citations.\n"
            "**If the user is setting their name, extract ONLY the name and use `set_name` as the intent.**\n"
            "Example:\n"
            "User: 'my name is John'\n"
            "Response: [ { \"intent\": \"set_name\", \"value\": \"John\" } ]\n"
            "Only reply with the JSON array. No explanation."
            "**Only** use `generate_code` if the user is explicitly asking for code.\n"
        )

        try:
            response = openai.ChatCompletion.create(
                model="gpt-4o",
                messages=conversation_history + [{"role": "user", "content": intent_prompt}]
            )
            content = response['choices'][0]['message']['content'].strip()

            # Make sure the content looks like JSON before parsing
            if not content.startswith("[") or "intent" not in content:
                raise ValueError("Invalid intent structure")

            actions = json.loads(content)

        except Exception as e:
            self.log_message.emit(f"Intent parsing failed, falling back to GPT: {e}")
            return get_gpt_response(command)

        # Execute each action in order
        final_response = ""
        for action in actions:
            intent = action.get("intent")
            value = action.get("value", "")

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
                final_response += get_gpt_response(value) + "\n"

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
        return final_response.strip()

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
SEARCH_WORD = "search for"
SCREEN_WORD = "screen"
sleep_mode = True


def get_serial_number():
    try:
        c = wmi.WMI()
        for bios in c.Win32_BIOS():
            return bios.SerialNumber.strip()
    except Exception as e:
        print("Failed to get serial number:", e)
        return None

def get_user_name(serial_number):
    try:
        conn = mysql.connector.connect(
            host="localhost",
            user="root",
            password="a1h2m3e4d5",
            database="chatbot_user_identification"
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
    try:
        conn = mysql.connector.connect(
            host="localhost",
            user="root",
            password="a1h2m3e4d5",
            database="chatbot_user_identification"
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

USER_SERIAL = get_serial_number()
USER_NAME = get_user_name(USER_SERIAL)

def speak_text(text, rate=200):
    """Convert text to speech with user's name if available."""
    engine.setProperty("rate", rate)
    if USER_NAME:
        text = text.replace("Joe AI:", f"Joe AI: {USER_NAME},")
    engine.say(text)
    engine.runAndWait()

def clean_markdown(text):
    """Removes Markdown symbols for clean speech output."""
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

def transcribe_speech(self, timeout=10, retries=2):
    recognizer = sr.Recognizer()
    recognizer.pause_threshold = 1.5
    try:
        with sr.Microphone() as source:
            recognizer.adjust_for_ambient_noise(source, duration=0.5)
            for attempt in range(retries):
                try:
                    self.log_message.emit("Listening...")
                    audio = recognizer.listen(source, timeout=timeout, phrase_time_limit=15)
                    text = recognizer.recognize_google(audio).lower().strip()
                    self.log_message.emit(f"You said: {text}")
                    return text
                except sr.WaitTimeoutError:
                    self.log_message.emit("Listening timed out.")
                except sr.UnknownValueError:
                    self.log_message.emit("Didn't catch that. Listening again...")
    except Exception as e:
        self.log_message.emit(f"Mic error: {e}")
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
    """Captures the screen, extracts visible text/icons, and answers user queries based on the screen details."""
    try:
        # Capture the screen
        screenshot_path = capture_screenshot()
        screen_image = cv2.imread(screenshot_path)

        # Step 1: Extract All Visible Text Using OCR
        extracted_text = pytesseract.image_to_string(screen_image, lang="eng").strip()

        # Step 2: Identify UI Elements (Icons, Buttons, Menus)
        detected_elements = detect_ui_elements(screen_image)

        # Step 3: Generate a Structured Screen Description
        screen_analysis = (
            f" **Extracted Text:**\n{extracted_text if extracted_text else 'No text detected.'}\n\n"
            f" **Detected UI Elements:**\n{', '.join(detected_elements) if detected_elements else 'No icons or buttons detected.'}"
        )

        # Remove screenshot after processing
        os.remove(screenshot_path)

        # Step 4: Pass Only the User Query and Screen Analysis to GPT-4
        combined_query = (
            f"The user is asking: '{user_query}'. "
            f"Here is everything visible on the screen:\n{screen_analysis}. "
            f"Provide a response ONLY based on the user's question and the extracted screen content."
        )

        # Get GPT-4's response
        final_response = get_gpt_response(combined_query)

        return final_response  # Return AI's response based on extracted details

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

def get_gpt_response(user_query):
    try:
        conversation_history.append({"role": "user", "content": user_query})
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=conversation_history
        )
        gpt_response = response['choices'][0]['message']['content']
        plain_text_response = clean_markdown(gpt_response)

        # Add user's name if it exists
        if USER_NAME:
            plain_text_response = f"{USER_NAME}, {plain_text_response}"

        print(f"Joe AI: {plain_text_response}")
        speak_text(plain_text_response)
        return plain_text_response
    except Exception as e:
        print(f"Error communicating with GPT: {e}")
        return "I encountered an issue retrieving an answer."


class JoeAIApp(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.listener_thread = VoiceListener()
        self.listener_thread.speech_detected.connect(self.update_chat)
        self.listener_thread.user_speaking.connect(self.show_user_animation)
        self.listener_thread.ai_speaking.connect(self.show_ai_animation)
        self.listener_thread.log_message.connect(self.update_chat)
        self.listener_thread.ai_response.connect(self.update_chat)
        self.listener_thread.start()

    def init_ui(self):
        self.setWindowTitle("Joe AI")
        self.setGeometry(100, 100, 600, 700)
        self.layout = QVBoxLayout()

        self.chat_display = QTextEdit()
        self.chat_display.setReadOnly(True)
        self.layout.addWidget(self.chat_display)

        self.animation_layout = QHBoxLayout()
        self.user_animation = QLabel(self)
        self.user_animation.setAlignment(Qt.AlignCenter)
        self.animation_layout.addWidget(self.user_animation)
        self.user_movie = QMovie("user_speaking.gif")
        self.user_animation.setMovie(self.user_movie)
        self.user_animation.setFixedSize(250, 250)
        
        self.ai_animation = QLabel(self)
        self.ai_animation.setAlignment(Qt.AlignCenter)
        self.animation_layout.addWidget(self.ai_animation)
        self.ai_movie = QMovie("ai_speaking.gif")
        self.ai_animation.setMovie(self.ai_movie)
        self.ai_animation.setFixedSize(250, 250)

        self.layout.addLayout(self.animation_layout)

        self.close_button = QPushButton("Close Window")
        self.close_button.clicked.connect(self.hide)
        self.layout.addWidget(self.close_button)

        self.setLayout(self.layout)
        self.tray_icon = SystemTrayApp(self)

    def update_chat(self, text):
        self.chat_display.append(text)
        self.user_movie.stop()
        self.ai_movie.stop()

    def show_user_animation(self, status):
        if status:
            self.user_movie.start()
            self.ai_animation.hide()
        else:
            self.user_movie.stop()
            self.ai_animation.show()

    def show_ai_animation(self, status):
        if status:
            self.ai_movie.start()
            self.user_animation.hide()
        else:
            self.ai_movie.stop()
            self.user_animation.show()

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
