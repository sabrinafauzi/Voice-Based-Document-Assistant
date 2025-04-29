import os
import sys
import json
import requests
import pyttsx3
import sounddevice as sd
import numpy as np
import keyboard
import fitz  # PyMuPDF
import docx  # python-docx
from vosk import Model, KaldiRecognizer
import soundfile as sf

interrupted = False

def stop_interruption(event):
    global interrupted
    if event.name == "esc":
        interrupted = True
        print("Interruption detected! Exiting program...")
        engine.stop()
        sys.exit()

keyboard.on_press_key("esc", stop_interruption)

# Initialize TTS with pyttsx3
engine = pyttsx3.init()
engine.setProperty('rate', 150)
engine.setProperty('volume', 1)

def speak_text(text):
    global interrupted
    if interrupted:
        sys.exit()
    engine.say(text)
    engine.runAndWait()
   
def play_sound(file_path):
    try:
        data, fs = sf.read(file_path)
        sd.play(data, fs)
        sd.wait()
    except Exception as e:
        print(f"Error playing sound: {e}")
        speak_text("Error playing sound.")

# Initialize STT with Vosk
def get_microphone_device():
    try:
        devices = sd.query_devices()
        for i, device in enumerate(devices):
            if device['max_input_channels'] > 0:
                return i
    except Exception as e:
        print(f"Error detecting microphone: {e}")
        speak_text("Error detecting microphone. Please check your device.")
    return None

def transcribe_audio_from_microphone(duration=10):
    model_path = r"C:\Voice-Based-Document-Assistant\vosk-model-small-en-us-0.15"
    if not os.path.exists(model_path):
        print("Speech model not found. Please check the file path.")
        speak_text("Speech model not found. Please check the file path.")
        return ""
   
    model = Model(model_path)
    recognizer = KaldiRecognizer(model, 16000)
    fs = 16000
    device = get_microphone_device()

    if interrupted:
        return ""

    if device is None:
        print("No valid microphone detected!")
        speak_text("No valid microphone detected!")
        return ""

    for _ in range(3):
        play_sound("C:\Voice-Based-Document-Assistant\mic.wav") #Adjust Path
        audio_data = sd.rec(int(duration * fs), device=device, samplerate=fs, channels=1, dtype='int16')
        sd.wait()

        if interrupted:
            return ""
               
        if recognizer.AcceptWaveform(audio_data.tobytes()):
            result = json.loads(recognizer.Result())
            transcribed_text = result.get("text", "").strip()
            if transcribed_text:
                print(f"You said: {transcribed_text}. Is this correct? Press 'j' for yes or 'f' for no.")
                speak_text(f"You said: {transcribed_text}. Is this correct? Press 'j' for yes or 'f' for no.")
                while True:
                    if keyboard.is_pressed("j"):
                        return transcribed_text
                    elif keyboard.is_pressed("f"):
                        break
        print("Please try again.")
        speak_text("Please try again.")
    return ""

def find_document_files(keyword, search_path="C:/Voice-Based-Document-Assistant/Article"): #Adjust Path Document Folder
    matches = []
    for root, dirs, files in os.walk(search_path):
        for file in files:
            if file.lower() == f"{keyword.lower()}.pdf" or file.lower() == f"{keyword.lower()}.docx":
                matches.append(os.path.join(root, file))
    return matches

def read_pdf_by_page(file_path):
    doc = fitz.open(file_path)
    return [page.get_text("text") for page in doc]

def read_docx(file_path, paragraphs_per_page=5):
    doc = docx.Document(file_path)
    paragraphs = [para.text for para in doc.paragraphs if para.text.strip()]
    return [" ".join(paragraphs[i:i+paragraphs_per_page]) for i in range(0, len(paragraphs), paragraphs_per_page)]

def summarize_document(text):
    summary_prompt = "Summarize this document briefly:"
    return send_to_lmstudio(summary_prompt, text)

def read_aloud(pages):
    if not pages:
        print("This document is empty or could not be read.")
        speak_text("This document is empty or could not be read.")
        return
   
    current_page = 0
    total_pages = len(pages)
   
    while current_page < total_pages:
        if interrupted:
            sys.exit()
       
        speak_text(f"Reading page {current_page + 1}")
        speak_text(pages[current_page])

        if interrupted:
            sys.exit()
       
        print("Press 'j' to read the next page, press 'f' to read the previous page, or press 'space' to stop reading.")
        speak_text("Press 'j' to read the next page, press 'f' to read the previous page, or press 'space' to stop reading.")
        while True:
            if interrupted:
                sys.exit()
            if keyboard.is_pressed("j"):
                current_page += 1
                break
            elif keyboard.is_pressed("f") and current_page > 0:
                current_page -= 1
                break
            elif keyboard.is_pressed("space"):
                speak_text("Stopping document reading.")
                return

def send_to_lmstudio(transcribed_text, document_text, conversation_history=None):
    if interrupted:
        sys.exit()

    if conversation_history is None:
        conversation_history = []

    url = "http://127.0.0.1:1234/v1/chat/completions"
    headers = {"Content-Type": "application/json"}
    conversation_history.append({"role": "user", "content": transcribed_text})
    payload = {
        "messages": [{"role": "system", "content": "Answer only based on the following document. If the question is not relevant to the document, state that you can only answer questions related to this document."},
                      {"role": "user", "content": document_text[:4000]}] + conversation_history,
        "model": "mistral-7b-instruct-v0.1",
        "temperature": 0.7,
        "max_tokens": 100,
        "stream": False
    }
    try:
        if interrupted:
            sys.exit()
           
        print("Please wait, the program is generating the answer...")
        speak_text("Please wait, the program is generating the answer...")
        response = requests.post(url, json=payload, headers=headers)

        if interrupted:
            sys.exit()

        response.raise_for_status()
        data = response.json()
        chatbot_response = data.get("choices", [{}])[0].get("message", {}).get("content", "").strip()

        if chatbot_response:
            print(chatbot_response)
            speak_text(chatbot_response)
            return chatbot_response, conversation_history
        print("No response from chatbot.")
        speak_text("No response from chatbot.")

    except requests.exceptions.RequestException as e:
        print(f"Error connecting to chatbot: {e}")
        speak_text(f"Error connecting to chatbot")

    return "", conversation_history

def menu_options():
    print("")
    print("Please choose an option:")
    print("1. Choose a new file")
    print("2. Read a file")
    print("3. File summary")
    print("4. Ask a question")
    print("5. Exit program")
    print("Say the option number you want to choose.")

    speak_text("Please choose an option: One for choose a new file, Two for read a file, Three for file summary, Four for ask a question, Five to exit the program. Say the option number you want to choose.")
    valid_inputs = {"one": 1, "two": 2, "three": 3, "for": 4, "five": 5}

    while True:
        choice = transcribe_audio_from_microphone(duration=5).lower()
        if not choice:
            continue
        if choice in valid_inputs:
            return valid_inputs[choice]
        else:
            print("Invalid choice, please say again.")
            speak_text("Invalid choice, please say again.")
            continue

def start_program():
    while True:
        print("Please say the file name you want to open")
        speak_text("Please say the file name you want to open")
        transcribed_text = transcribe_audio_from_microphone(duration=5)
        if interrupted:
            break
        if transcribed_text:
            files = find_document_files(transcribed_text)
            if not files:
                print("File not found. Please try again.")
                speak_text("File not found. Please try again.")
                continue
            elif len(files) > 1:
                print("Multiple files found. Use the 'f' key to move to the previous file, the 'j' key to move to the next file, and press the 'space' key to select the file.")
                speak_text("Multiple files found. Use the 'f' key to move to the previous file, the 'j' key to move to the next file, and press the 'space' key to select the file.")
                options = [f for f in files if f.endswith(".pdf") or f.endswith(".docx")]
                selected_index = 0
               
                while True:
                    print(f"Currently selected: {os.path.basename(options[selected_index])}")
                    speak_text(f"Currently selected: {os.path.basename(options[selected_index])}")
   
                    while True:
                        if keyboard.is_pressed("f"):
                            selected_index = (selected_index - 1) % len(options)
                            break
                        elif keyboard.is_pressed("j"):
                            selected_index = (selected_index + 1) % len(options)
                            break
                        elif keyboard.is_pressed("space"):
                            file_path = options[selected_index]
                            break
                    if keyboard.is_pressed("space"):  
                        break
            else:
                file_path = files[0]
           
            speak_text(f"Opening {os.path.basename(file_path)} file")
            document_pages = read_pdf_by_page(file_path) if file_path.endswith(".pdf") else read_docx(file_path)

            while True:
                choice = menu_options()
                if choice == 1:
                    print("Selected: Choose a new file")
                    speak_text("Selected: Choose a new file")
                    break
                elif choice == 2:
                    print("Selected: Read a file")
                    speak_text("Selected: Read a file")
                    read_aloud(document_pages)
                elif choice == 3:
                    print("Selected: File summary")
                    speak_text("Selected: File summary")
                    summarize_document(" ".join(document_pages))
                elif choice == 4:
                    print("Selected: Ask a question")
                    speak_text("Selected: Ask a question")
                    print("What do you want to ask about the document.")
                    speak_text("What do you want to ask about the document.")
                    document_text = " ".join(document_pages)

                    while True:
                        transcribed_text = transcribe_audio_from_microphone()
                        if transcribed_text:
                            send_to_lmstudio(transcribed_text, document_text)
                            break
                elif choice == 5:
                    print("Exiting program.")
                    speak_text("Exiting program.")
                    sys.exit()

if __name__ == "__main__":
    print("Press 'Space' to start the program")
    speak_text("Press 'Space' to start the program")
    while True:
        if keyboard.is_pressed("space"):
            print("Starting the program")
            start_program()
            break
