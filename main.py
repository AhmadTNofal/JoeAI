import openai
import speech_recognition as sr
import pyttsx3

# Set your OpenAI API key
openai.api_key = "sk-tbtbdjXN0PNeKVX8x6oXJFABUkwYsEeOj9TinWn3jOT3BlbkFJuGto6skfATpazIFkDBnEr1JtKDe0ykgJkavseRQP0A"

def transcribe_speech(timeout=10):
    """
    Captures speech from the default microphone and converts it to text.
    Handles timeout and other exceptions gracefully.
    :param timeout: Time to wait for the user to start speaking.
    :return: Transcribed text or None if no speech was detected.
    """
    recognizer = sr.Recognizer()

    try:
        with sr.Microphone() as source:  # Use default microphone
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
    Converts text to speech using pyttsx3 with adjustable rate.
    :param text: Text to convert to speech.
    :param rate: Speech rate (default is 200 words per minute).
    """
    engine = pyttsx3.init()
    engine.setProperty('rate', rate)  # Set the speech rate
    engine.say(text)
    engine.runAndWait()

def main():
    print("Welcome! Speak into your microphone, and I will respond.")
    print("Say 'exit' to quit.")

    # Initialize conversation history with a system message
    conversation_history = [
        {"role": "system", "content": "You are a helpful assistant."}
    ]

    while True:
        # Get speech input
        user_input = transcribe_speech(timeout=10)
        if user_input:
            if user_input.lower() == "exit":
                print("Goodbye!")
                break

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
