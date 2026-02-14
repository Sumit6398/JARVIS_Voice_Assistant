''' This is a simple voice assistant program that listens to user commands and
performs various tasks such as opening
websites, playing music, telling the time, and more. 
It uses the SpeechRecognition library for voice recognition,
win32com for text-to-speech, and the Ollama API for AI responses.'''

# Import necessary libraries
import speech_recognition as sr             # For recognizing speech input from the user
import os                                   # For interacting with the operating system (e.g., opening files, applications)
import win32com.client                      # For text-to-speech functionality using Windows SAPI
import webbrowser                           # For opening websites in the default web browser
from datetime import datetime                             # For handling date and time operations
import ollama                               # For interacting with the Ollama API to get AI responses based on user queries


speaker = win32com.client.Dispatch("SAPI.SpVoice")

def speak(text):
    print("Jarvis:", text)
    speaker.Speak(text)

def save_response_to_file(response):
    """Save AI response with timestamp"""
    with open("jarvis_ai_answers.txt", "a", encoding="utf-8") as file:
        time_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        file.write(f"[{time_now}] {response}\n")


# Function to get AI response from Ollama API
def ai(prompt):
    try:
        response = ollama.chat(       
            model='phi',
            messages=[
                {"role": "user", "content": prompt}      # The user's query is sent as a message to the AI model
            ]
        )

        output = response['message']['content']
        print(output)
        speak(output) 
        save_response_to_file(output)
         
    except Exception as e:          # If there is an error while getting the AI response, it will be caught and printed, and a message will be spoken to the user indicating that something went wrong.
        print("Error:", e)
        speak("Sorry sir, something went wrong.")

recognizer = sr.Recognizer()       # Initialize the speech recognizer to listen for user commands through the microphone.


def takeCommand():                # This function listens to the user's voice input and converts it into text using Google's speech recognition service. It handles exceptions for unrecognized speech and network errors, providing appropriate feedback to the user.
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.adjust_for_ambient_noise(source, duration=1)   # Adjusts the recognizer sensitivity to ambient noise to improve recognition accuracy
        audio = recognizer.listen(source)
        # try:
        #     audio = recognizer.listen(source, timeout=5, phrase_time_limit=8)
        # except sr.WaitTimeoutError:
        #     print("Listening timed out.")
        #     return None


    try:
        print("Recognizing...")
        query = recognizer.recognize_google(audio, language='en-IN')    # Recognizes the speech using Google's speech recognition service, specifying the language as English (India)
        print("You said:", query)
        return query.lower()

    except sr.UnknownValueError:       # If the speech is unintelligible or cannot be recognized, this exception is caught, and a message is printed and spoken to the user indicating that the input was not understood.
        print("Sorry, I did not understand that.")
        return None

    except sr.RequestError:
        print("Network error.")
        return None


# Greet once
speak("Hello, I am Jarvis AI, HOW CAN I HELP YOU SIR?")

# Main Loop
while True:
    query = takeCommand()     # The program enters an infinite loop where it continuously listens for user commands, processes them, and performs actions based on the recognized commands. The loop will keep running until the program is manually stopped or an exit command is given.
    
    if query is None:
        continue
    # Open websites which are mentioned in the list
    # todo: Add more sites
    
    
    '''The program checks if the user's query contains a command to open a specific website. It iterates through a predefined list
    of websites, and if the query matches the format "Open [website name]", it will open the corresponding website in the default 
    web browser and speak a confirmation message to the user.'''
    
    
    sites = [["youtube","https://www.youtube.com"], 
             ["google","https://www.google.com"],
             ["wikipedia","https://www.wikipedia.org"],
             ["twitter","https://www.twitter.com"],
             ["gen spark","https://www.genspark.ai"],
             ["github","https://www.github.com"],
             ["linkedin","https://www.linkedin.com"]]
    for site in sites:                               
        if f"Open {site[0]}".lower() in query.lower():
            speak(f"Opening {site[0]} sir...")
            webbrowser.open(site[1])
            
    # opening the music file
    # todo: Add a feature to play a specific song
    if "open music" in query:
        musicPath = r"C:/Users/sumit/Music/Meri_Zindagi_Hai_Tu_1.mp3"
        os.startfile(musicPath)
    
    # telling the current time
    
    if "the time" in query:

        strTime = datetime.datetime.now().strftime("%H:%M:%S")
        speak(f"Sir, the time is {strTime}")
        
    # Greet based on time
    
        hour = datetime.datetime.now().hour
        if hour >= 0 and hour < 12:       
            speak("Good Morning Sir!")
        elif hour >= 12 and hour < 18:
            speak("Good Afternoon Sir!")    
        else:
            speak("Good Evening Sir!")
            
    # giving the command to open excel
    
    if "open excel" in query:
        excelpath = r"C:/ProgramData/Microsoft/Windows/Start Menu/Programs/Excel.lnk" 
        os.startfile(excelpath)
    
    # giving the command to open word
    
    if "open word" in query:
        wordpath = r"C:/ProgramData/Microsoft/Windows/Start Menu/Programs/Word.lnk" 
        os.startfile(wordpath)

    # giving the command to open powerpoint
    
    if "open powerpoint" in query:
        powerpointpath = r"C:/ProgramData/Microsoft/Windows/Start Menu/Programs/PowerPoint.lnk"
        os.startfile(powerpointpath)
        
    # giving the command to open camera application 
      
    if "open camera" in query:
        cameraPath = r"microsoft.windows.camera:"
        webbrowser.open(cameraPath)
        
    # giving command to greet vallentine's day
    
    if "valentine's day" in query:
        speak("Happy Valentine's Day to you sir, I hope you have a wonderful day filled with love and joy!  nandini is the best choice for you ")
        search_query = query.replace("valentine's day", "")
        webbrowser.open(f"https://www.genspark.ai/api/code_sandbox_light/preview/"f"a329d76b-cb95-44d0-8c20-d20f895dfed0/index.html"f"?canvas_history_id=63c9d9e4-0092-4e0d-8a99-427c523b48b0&query={search_query}")
        
    if "using source".lower() in query.lower():  # If the user's query contains the phrase "Using artificial intelligence", the program will respond by speaking a message that highlights the capabilities of the AI assistant, such as its ability to perform various tasks and provide information. Additionally, it will open a specific URL in the web browser, which appears to be a page related to using artificial intelligence, possibly providing more information or resources on the topic.
        ai(prompt=query)
    
    if "stop" in query or "abort" in query:
        stop_speaking = True
        speaker.Speak("", 3)
        continue

    if "exit" in query :
        speak("Goodbye sir.")
        break    
    speak(query)




