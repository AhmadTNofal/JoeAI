import openai
import pyttsx3
import pyautogui
import pytesseract
import cv2
import numpy as np
from PIL import Image
import speech_recognition as sr

# Set your OpenAI API key
openai.api_key = "sk-tbtbdjXN0PNeKVX8x6oXJFABUkwYsEeOj9TinWn3jOT3BlbkFJuGto6skfATpazIFkDBnEr1JtKDe0ykgJkavseRQP0A"

# Initialize text-to-speech engine
engine = pyttsx3.init()

def transcribe_speech(timeout=10):
    """
    Captures speech from the default microphone and converts it to text.
    Handles timeout and other exceptions gracefully.
    :param timeout: Time to wait for the user to start speaking.
    :return: Transcribed text or None if no speech was detected.
    """
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
    """
    Captures the entire screen and saves it as an image.
    :return: PIL Image object of the screen.
    """
    screenshot = pyautogui.screenshot()
    return screenshot

def extract_text_from_image(image):
    """
    Extracts text from a given image using Tesseract OCR.
    :param image: PIL Image object.
    :return: Extracted text as a string.
    """
    try:
        text = pytesseract.image_to_string(image)
        return text.strip()
    except Exception as e:
        print(f"Error in text extraction: {e}")
        return None

def detect_objects(image):
    """
    Detects objects (like icons or windows) in the given image using OpenCV and YOLOv5.
    :param image: PIL Image object.
    :return: List of detected objects with labels.
    """
    # Convert PIL Image to NumPy array
    image_np = np.array(image)

    # Placeholder for object detection logic
    # Replace this with real detection logic using a pre-trained model
    detected_objects = ["desktop icon", "desktop icon", "taskbar", "open application window"]
    return detected_objects

def interpret_screen(user_command, conversation_history):
    """
    Captures the screen, extracts data, and interprets the screen dynamically based on the user's command.
    :param user_command: The text command from the user.
    :param conversation_history: List of messages forming the conversation.
    """
    print("Capturing screen...")
    screen_image = capture_screen()

    # Extract text and detect objects on the screen
    print("Extracting text from the screen...")
    screen_text = extract_text_from_image(screen_image)

    print("Detecting objects on the screen...")
    detected_objects = detect_objects(screen_image)

    # Create a detailed screen description
    screen_data = "Here is what I see on the screen:\n"
    if screen_text:
        screen_data += f"- Extracted text: {screen_text}\n"
    if detected_objects:
        screen_data += f"- Detected objects: {', '.join(detected_objects)}\n"

    print(f"Screen Data: {screen_data}")

    # Generate a GPT prompt with clear instructions
    prompt = (
        f"User asked: \"{user_command}\". Based on the screen data provided below, "
        f"analyze and answer the user's query as best as possible:\n{screen_data}\n"
        "Only use the provided screen data to answer the question."
    )

    print("Sending prompt to GPT...")

    # Add screen context and user command to conversation history
    conversation_history.append({"role": "user", "content": user_command})
    conversation_history.append({"role": "assistant", "content": screen_data})

    # Get GPT's response
    gpt_response = get_gpt_response(conversation_history)
    if gpt_response:
        print(f"GPT-4 Response: {gpt_response}")
        # Add GPT response to conversation history
        conversation_history.append({"role": "assistant", "content": gpt_response})
        # Speak the response
        speak_text(gpt_response)
    else:
        print("Failed to get a response from GPT.")

def get_gpt_response(conversation_history):
    """
    Sends the conversation history to GPT and gets a response.
    :param conversation_history: List of messages forming the conversation.
    :return: GPT's response text or None if an error occurs.
    """
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=conversation_history
        )
        return response['choices'][0]['message']['content']
    except Exception as e:
        print(f"Error communicating with GPT: {e}")
        return None

def speak_text(text, rate=200):
    """
    Converts text to speech using pyttsx3.
    :param text: Text to convert to speech.
    """
    engine.setProperty('rate', rate)  # Set the speech rate
    engine.say(text)
    engine.runAndWait()

def main():
    print("Welcome! Speak into your microphone, and I will respond.")
    print("Say 'exit' to quit or ask about your screen (e.g., 'How many icons are on my screen?').")


    # Initialize conversation history with a system message
    conversation_history = [
        {"role": "system", "content": "You are a helpful assistant that can analyze the user's screen and answer questions dynamically."}
    ]

    while True:
        # Get speech input
        user_input = transcribe_speech(timeout=10)
        if user_input:
            if user_input.lower() == "exit":
                print("Goodbye!")
                break
            elif "screen" in user_input.lower():
                interpret_screen(user_input, conversation_history)
            else:
                # Add user input to conversation history
                conversation_history.append({"role": "user", "content": user_input})

                # Get GPT response
                response = get_gpt_response(conversation_history)
                if response:
                    print(f"GPT-4: {response}")
                    # Add GPT response to conversation history
                    conversation_history.append({"role": "assistant", "content": response})
                    # Speak the response
                    speak_text(response)
                else:
                    print("Failed to get a response from GPT.")
        else:
            print("No input detected. Try again.")

if __name__ == "__main__":
    main()
