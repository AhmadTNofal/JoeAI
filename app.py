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

# Set your OpenAI API key
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
                    self.speech_detected.emit("I'm listening!")
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
            "Your task is to extract user intent from this request.\n"
            "Return ONLY a JSON array of objects, like this:\n"
            "[\n"
            "  {{ \"intent\": \"open_app\", \"value\": \"Spotify\" }},\n"
            "  {{ \"intent\": \"web_search\", \"value\": \"lofi music\" }}\n"
            "]\n\n"
            "Valid intents:\n"
            "- general_chat\n"
            "- web_search\n"
            "- open_app\n"
            "- close_app\n"
            "- analyze_screen\n"
            "- sleep\n"
            "- exit\n\n"
            "**DO NOT** explain your answer, greet the user, or say anything else — ONLY respond with valid JSON.\n"
            "If it's just a general chat message, return:\n"
            "[ {{ \"intent\": \"general_chat\", \"value\": \"original user message\" }} ]\n"
        )

        try:
            response = openai.ChatCompletion.create(
                model="gpt-4o",
                messages=conversation_history + [{"role": "user", "content": intent_prompt}]
            )
            import json
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

def speak_text(text, rate=200):
    """Convert text to speech."""
    engine.setProperty("rate", rate)
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

def get_gpt_response(user_query):
    """Handles general queries using GPT-4 and cleans Markdown formatting for TTS."""
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
    app = QApplication(sys.argv)
    joe_ai = JoeAIApp()
    joe_ai.show()
    sys.exit(app.exec_())
