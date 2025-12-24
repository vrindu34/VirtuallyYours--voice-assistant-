import subprocess
import wolframalpha
import pyttsx3
import json
import speech_recognition as sr
import datetime
import openai
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import getpass
import wmi
from pathlib import Path
from bs4 import BeautifulSoup
import urllib.parse
import random
import sys


from dotenv import load_dotenv
import os

load_dotenv()

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
WOLFRAM_ALPHA_APP_ID = os.getenv("WOLFRAM_ALPHA_APP_ID")
WEATHER_API_KEY = os.getenv("WEATHER_API_KEY")
NEWS_API_KEY = os.getenv("NEWS_API_KEY")

# Initialize OpenAI
openai.api_key = OPENAI_API_KEY

# ========== ASSISTANT NAME CONFIGURATION ==========
ASSISTANT_NAME = None  # Will be set by user

# ==========  SAPI SPEECH ENGINE ==========
def speak(text):
    """Speak text using Windows SAPI"""
    if not text or text == "None":
        return
    
    print(f"{ASSISTANT_NAME}: {text}")
    
    try:
        # Clean the text for command line
        text_clean = text.replace('"', '\\"').replace("'", "''")
        
        # Use Windows SAPI through PowerShell
        ps_command = f'''powershell -Command "$voice = New-Object -ComObject SAPI.SpVoice; $voice.Speak('{text_clean}')"'''
        
        # Execute the command silently
        process = subprocess.Popen(ps_command, 
                                  shell=True, 
                                  stdout=subprocess.PIPE, 
                                  stderr=subprocess.PIPE)
        process.communicate(timeout=10)
        
    except subprocess.TimeoutExpired:
        process.kill()
        print("Speech timeout")
    except Exception as e:
        print(f"Speech error: {e}")
        print(f"[Text only]: {text}")

# ========== SPEECH RECOGNITION ==========
def takeCommand():
    """Take voice command from user"""
    r = sr.Recognizer()
    
    try:
        with sr.Microphone() as source:
            print("\nListening... Speak now")
            r.adjust_for_ambient_noise(source, duration=0.5)
            
            try:
                audio = r.listen(source, timeout=5, phrase_time_limit=5)
                print("Recognizing...")
                
                query = r.recognize_google(audio, language='en-in')
                print(f"You: {query}")
                return query.lower()
                
            except sr.WaitTimeoutError:
                print("Listening timeout")
                return None
            except sr.UnknownValueError:
                print("Could not understand audio")
                return None
            except sr.RequestError as e:
                print(f"Speech recognition error: {e}")
                return None
                
    except Exception as e:
        print(f"Microphone error: {e}")
        return None

def takeCommandname():
    """Take name input from user with better handling"""
    r = sr.Recognizer()
    
    with sr.Microphone() as source:
        print("Listening for your response...")
        
        # Adjust settings for better recognition
        r.energy_threshold = 300  # Lower threshold for better sensitivity
        r.dynamic_energy_threshold = True
        r.pause_threshold = 1.0  # Wait 1 second before considering speech ended
        
        try:
            # Listen with longer timeout and no phrase limit
            audio = r.listen(source, timeout=8, phrase_time_limit=None)
            print("Recognizing...")
            
            query = r.recognize_google(audio, language='en-in')
            print(f"You: {query}")
            return query
            
        except sr.WaitTimeoutError:
            print("No speech detected within timeout")
            return None
        except sr.UnknownValueError:
            print("Google Speech Recognition could not understand audio")
            return None
        except sr.RequestError as e:
            print(f"Could not request results from Google Speech Recognition service; {e}")
            return None
        except Exception as e:
            print(f"Error in speech recognition: {e}")
            return None

# ========== SET ASSISTANT NAME ==========
def set_assistant_name():
    """Let user set the assistant's name"""
    global ASSISTANT_NAME
    
 
    
    # Test speech first
    test_text = "Welcome to Virtually Yours setup!"
    print(f"System: {test_text}")
    try:
        text_clean = test_text.replace('"', '\\"').replace("'", "''")
        ps_command = f'''powershell -Command "$voice = New-Object -ComObject SAPI.SpVoice; $voice.Speak('{test_text}')"'''
        subprocess.run(ps_command, shell=True, capture_output=True, timeout=5)
        print("Speech system working!")
    except:
        print("Speech system may not be working properly")
    
    time.sleep(1)
    
    # Ask if user wants to name the assistant
    print("\nWould you like to give me a name?")
    print("You can say: Yes or choose yourself")
    
    # Speak the question
    question = "Would you like to give me a name? You can say Yes or choose yourself"
    speak(question)
    
    response = takeCommandname()
    
    if response:
        response = response.lower()
        
        # User wants to name the assistant
        if 'yes' in response or 'give me' in response or 'i want' in response:
            speak("What would you like to call me?")
            print("\nWhat would you like to call me?")
            
            name_response = takeCommandname()
            if name_response:
                ASSISTANT_NAME = name_response.strip()
                speak(f"Thank you! From now on, I'll be called {ASSISTANT_NAME}")
                print(f"\nAssistant name set to: {ASSISTANT_NAME}")
                return
        
        # User wants assistant to choose
        elif 'choose' in response or 'yourself' in response or 'pick' in response:
            names = ["Alexa", "Jarvis", "Siri", "Cortana", "Athena", "Orion", "Nova", "Zen", "Aura", "Echo"]
            chosen_name = random.choice(names)
            ASSISTANT_NAME = chosen_name
            speak(f"I've chosen the name {ASSISTANT_NAME}. I hope you like it!")
            print(f"\nAssistant name set to: {ASSISTANT_NAME}")
            return
    
    # Default name if no response or other cases
    ASSISTANT_NAME = "Virtually Yours"
    speak(f"I'll go by {ASSISTANT_NAME}. Nice to meet you!")
    print(f"\nAssistant name set to: {ASSISTANT_NAME}")

# ========== CORE FUNCTIONS ==========
def wishMe():
    """Greet the user based on time"""
    hour = datetime.datetime.now().hour
    if hour >= 0 and hour < 12:
        speak("Good Morning!")
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")
    
    speak(f"I am {ASSISTANT_NAME}, your personal assistant")
    speak("How can I help you today?")

def usrname():
    """Get user's name"""
    speak("What should I call you?")
    uname = takeCommandname()
    if uname:
        speak(f"Welcome {uname}")
        print(f"\nUser: {uname}")
        return uname
    speak("I will call you Friend")
    return "Friend"

def get_quote():
    """Get random inspirational quote"""
    try:
        response = requests.get("https://api.quotable.io/random", timeout=5)
        if response.status_code == 200:
            data = response.json()
            return f"{data['content']} - {data['author']}"
    except:
        pass
    
    # Fallback quotes
    fallback_quotes = [
        "The only way to do great work is to love what you do. - Steve Jobs",
        "The future belongs to those who believe in the beauty of their dreams. - Eleanor Roosevelt",
        "It always seems impossible until it's done. - Nelson Mandela"
    ]
    return random.choice(fallback_quotes)

# ========== CALCULATION WITH WOLFRAM ALPHA ==========
def calculate(query):
    """Perform calculations using Wolfram Alpha"""
    try:
        speak("Calculating...")
        
        # Initialize Wolfram Alpha client
        client = wolframalpha.Client(WOLFRAM_ALPHA_APP_ID)
        
        # Send query to Wolfram Alpha
        res = client.query(query)
        
        # Get the result
        answer = next(res.results).text
        speak(f"The answer is {answer}")
        print(f"Calculation Result: {answer}")
        
    except StopIteration:
        speak("Sorry, I couldn't calculate that. The query format might be incorrect.")
    except Exception as e:
        print(f"Calculation error: {e}")
        speak("Sorry, I couldn't perform the calculation.")

# ========== WEATHER FUNCTION ==========
def get_weather():
    """Get weather information for a city"""
    try:
        speak("Which city would you like the weather for?")
        city = takeCommand()
        
        if not city or city == "none":
            city = "Delhi"
            speak("Using default city: Delhi")
        
        # Construct API URL
        base_url = "http://api.openweathermap.org/data/2.5/weather?"
        complete_url = f"{base_url}appid={WEATHER_API_KEY}&q={city}&units=metric"
        
        # Make API request
        response = requests.get(complete_url)
        data = response.json()
        
        if data["cod"] != "404":
            # Extract weather data
            main_data = data["main"]
            weather_data = data["weather"][0]
            
            temperature = main_data["temp"]
            description = weather_data["description"]
            humidity = main_data["humidity"]
            
            weather_report = f"Current weather in {city}: Temperature is {temperature} degrees Celsius with {description}. Humidity is {humidity} percent."
            
            print(f"Weather Report: {weather_report}")
            speak(weather_report)
            
        else:
            speak(f"Sorry, I couldn't find weather information for {city}. Please check the city name.")
            
    except Exception as e:
        print(f"Weather error: {e}")
        speak("Sorry, I couldn't fetch the weather information.")

# ========== OPENAI FUNCTION ==========
def answer_question(query):
    """Answer questions using OpenAI GPT"""
    try:
        speak("Let me think about that...")
        
        # Create completion
        response = openai.Completion.create(
            engine="text-davinci-003",
            prompt=query,
            max_tokens=150,
            temperature=0.7,
            top_p=1.0,
            frequency_penalty=0.0,
            presence_penalty=0.0
        )
        
        # Extract and speak the answer
        answer = response.choices[0].text.strip()
        if answer:
            print(f"AI Answer: {answer}")
            speak(answer)
        else:
            speak("I don't have an answer for that question.")
            
    except Exception as e:
        print(f"OpenAI error: {e}")
        speak("Sorry, I couldn't process your question at the moment.")

# ========== NEWS FUNCTION ==========
def get_news():
    """Get latest news headlines"""
    try:
        # Use News API
        url = f"https://newsapi.org/v2/top-headlines?country=in&apiKey={NEWS_API_KEY}"
        response = requests.get(url)
        
        if response.status_code == 200:
            news_data = response.json()
            articles = news_data.get('articles', [])[:5]
            
            if articles:
                speak("Here are the top news headlines")
                for i, article in enumerate(articles, 1):
                    headline = article.get('title', 'No title')
                    print(f"{i}. {headline}")
                    
                    # Speak first 2 headlines only
                    if i <= 2:
                        speak(headline)
                        time.sleep(0.5)
            else:
                speak("No news articles found.")
        else:
            speak("Could not fetch news. Opening Google News.")
            webbrowser.open("https://news.google.com")
            
    except Exception as e:
        print(f"News error: {e}")
        speak("Opening Google News for latest updates.")
        webbrowser.open("https://news.google.com")

# ========== COUNTDOWN TIMER ==========
def countdown_timer():
    """Start a countdown timer"""
    try:
        speak("Enter countdown time in seconds")
        time_input = takeCommand()
        
        if time_input and any(char.isdigit() for char in time_input):
            seconds = int(''.join(filter(str.isdigit, time_input)))
            
            if seconds > 0:
                speak(f"Starting countdown from {seconds} seconds")
                
                for i in range(seconds, 0, -1):
                    print(f"Countdown: {i} seconds")
                    if i <= 5:
                        speak(str(i))
                    time.sleep(1)
                
                speak("Time's up!")
            else:
                speak("Please specify a positive number of seconds")
        else:
            speak("Please say a number like 10 seconds")
            
    except Exception as e:
        print(f"Countdown error: {e}")
        speak("Countdown failed")

# ========== NOTES FUNCTION ==========
def write_note():
    """Write a note to file"""
    try:
        speak("What should I write in the note?")
        note_content = takeCommand()
        
        if note_content and note_content != "none":
            with open('assistant_notes.txt', 'a') as f:
                timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                f.write(f"{timestamp}: {note_content}\n")
            
            speak("Note saved successfully")
        else:
            speak("No note content received")
            
    except Exception as e:
        print(f"Write note error: {e}")
        speak("Could not save the note")

def read_notes():
    """Read notes from file"""
    try:
        if os.path.exists('assistant_notes.txt'):
            with open('assistant_notes.txt', 'r') as f:
                notes = f.read()
            
            if notes:
                speak("Here are your notes")
                print("\nYour Notes:")
                print(notes)
            else:
                speak("You have no notes saved")
        else:
            speak("No notes file found. You can create one by saying 'write note'")
            
    except Exception as e:
        print(f"Read notes error: {e}")
        speak("Could not read notes")

# ========== MAIN PROGRAM ==========
def main():
    # Clear screen
    os.system('cls')
    
    print("\n" + "="*70)
    print("VIRTUAL ASSISTANT WITH CUSTOM NAME")
    print("="*70)
    
    # First, set the assistant's name
    set_assistant_name()
    time.sleep(1)
    
    # Now greet user with the assistant's name
    wishMe()
    time.sleep(1)
    
    # Get user name
    username = usrname()
    time.sleep(0.5)
    
    # Offer daily quote
    speak(f"{username}, would you like to hear a motivational quote?")
    response = takeCommand()
    
    if response and ('yes' in response or 'sure' in response or 'please' in response):
        quote = get_quote()
        speak("Here is today's quote:")
        print(f"\nDaily Quote: {quote}")
        speak(quote)
    else:
        speak("Alright, let's continue.")
    
    speak(f"How can I assist you today, {username}?")
    
    # Main command loop
    while True:
        query = takeCommand()
        
        if not query:
            print("Waiting for command...")
            continue
        
        query = query.lower()
        
        # Exit commands
        if any(word in query for word in ['exit', 'quit', 'bye', 'goodbye', 'stop']):
            speak(f"Goodbye {username}! It was nice assisting you.")
            break
        
        # Change my name command
        elif 'change your name' in query or 'change name' in query or 'rename' in query:
            speak("What would you like to call me instead?")
            new_name = takeCommand()
            if new_name and new_name != "none":
                global ASSISTANT_NAME
                old_name = ASSISTANT_NAME
                ASSISTANT_NAME = new_name.strip()
                speak(f"Thank you! I've changed my name from {old_name} to {ASSISTANT_NAME}")
            else:
                speak("I didn't catch the new name.")
        
        # What is your name
        elif 'your name' in query and 'change' not in query:
            speak(f"My name is {ASSISTANT_NAME}. How can I help you?")
        
        # Who are you
        elif 'who are you' in query:
            speak(f"I am {ASSISTANT_NAME}, your personal virtual assistant created to help you with various tasks.")
        
        # Wikipedia search
        elif 'wikipedia' in query:
            search_query = query.replace("wikipedia", "").strip()
            if search_query:
                try:
                    speak(f"Searching Wikipedia for {search_query}")
                    results = wikipedia.summary(search_query, sentences=2)
                    speak("According to Wikipedia")
                    print(f"\nWikipedia Result: {results}")
                    speak(results)
                except:
                    speak("Sorry, I couldn't find that on Wikipedia.")
        
        # Google search
        elif 'google' in query and 'search' in query:
            search_term = query.replace("google", "").replace("search", "").strip()
            if search_term:
                speak(f"Searching Google for {search_term}")
                webbrowser.open(f"https://www.google.com/search?q={urllib.parse.quote(search_term)}")
        
        # Open websites
        elif 'open youtube' in query:
            speak("Opening YouTube")
            webbrowser.open("https://youtube.com")
        
        elif 'open google' in query:
            speak("Opening Google")
            webbrowser.open("https://google.com")
        
        elif 'open gmail' in query:
            speak("Opening Gmail")
            webbrowser.open("https://mail.google.com")
        
        # Play music
        elif 'play music' in query or 'play song' in query:
            speak("Playing music on YouTube")
            webbrowser.open("https://music.youtube.com")
        
        # Time
        elif 'time' in query:
            current_time = datetime.datetime.now().strftime("%I:%M %p")
            speak(f"The current time is {current_time}")
            print(f"Current Time: {current_time}")
        
        # Date
        elif 'date' in query:
            current_date = datetime.datetime.now().strftime("%B %d, %Y")
            speak(f"Today's date is {current_date}")
            print(f"Today's Date: {current_date}")
        
        # Joke
        elif 'joke' in query:
            try:
                joke = pyjokes.get_joke()
                speak("Here's a joke for you")
                print(f"Joke: {joke}")
                speak(joke)
            except:
                speak("Why was the computer cold? It left its Windows open!")
        
        # Calculation
        elif 'calculate' in query or ('what is' in query and any(op in query for op in ['+', '-', '*', '/'])):
            calculate(query)
        
        # Weather
        elif 'weather' in query:
            get_weather()
        
        # News
        elif 'news' in query:
            get_news()
        
        # System commands
        elif 'lock' in query and 'computer' in query:
            speak("Locking your computer")
            ctypes.windll.user32.LockWorkStation()
        
        elif 'shutdown' in query and 'computer' in query:
            speak("Shutting down computer in 30 seconds")
            os.system("shutdown /s /t 30")
        
        elif 'restart' in query and 'computer' in query:
            speak("Restarting computer in 30 seconds")
            os.system("shutdown /r /t 30")
        
        elif 'empty recycle bin' in query:
            speak("Emptying recycle bin")
            winshell.recycle_bin().empty(confirm=False)
            speak("Recycle bin emptied successfully")
        
        # Countdown timer
        elif 'countdown' in query or 'timer' in query:
            countdown_timer()
        
        # Notes
        elif 'write note' in query or 'take note' in query:
            write_note()
        
        elif 'read note' in query or 'read notes' in query:
            read_notes()
        
        # Answer questions
        elif 'what is' in query or 'who is' in query or 'how to' in query:
            answer_question(query)
        
        # Personal interaction
        elif 'how are you' in query:
            responses = [
                f"I'm doing great, thank you for asking {username}!",
                f"I'm fantastic! Ready to help you, {username}.",
                f"I'm good, thank you! How are you today, {username}?",
                f"I'm excellent! What can I do for you, {username}?"
            ]
            response = random.choice(responses)
            speak(response)
        
        # Location search
        elif 'where is' in query:
            location = query.replace("where is", "").strip()
            if location:
                speak(f"Searching for {location}")
                webbrowser.open(f"https://maps.google.com?q={urllib.parse.quote(location)}")
        
        # Open camera
        elif 'open camera' in query:
            speak("Opening camera")
            os.system('start microsoft.windows.camera:')
        
        # Help
        elif 'help' in query or 'what can you do' in query:
            speak(f"I can help you with many things, {username}. Here are some commands:")
            print("""
            Available Commands:
            1. 'time' - Get current time
            2. 'date' - Get today's date
            3. 'joke' - Tell a joke
            4. 'calculate [expression]' - Perform calculation
            5. 'weather' - Get weather for a city
            6. 'news' - Get latest news
            7. 'open youtube/google/gmail' - Open websites
            8. 'play music' - Play music on YouTube
            9. 'wikipedia [topic]' - Search Wikipedia
            10. 'lock computer' - Lock your PC
            11. 'shutdown computer' - Shutdown PC
            12. 'restart computer' - Restart PC
            13. 'empty recycle bin' - Empty recycle bin
            14. 'countdown' - Start a countdown timer
            15. 'write note' - Write a note
            16. 'read note' - Read your notes
            17. 'where is [location]' - Search location on maps
            18. 'open camera' - Open camera app
            19. 'change your name' - Change my name
            20. 'exit' or 'bye' - Close the assistant
            """)
            speak("You can say any of these commands to get started.")
        
        # Default response
        else:
            speak("I didn't understand that command. You can say 'help' to see what I can do.")

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nAssistant stopped by user.")
        # Try to speak goodbye
        try:
            text_clean = "Goodbye!".replace('"', '\\"').replace("'", "''")
            ps_command = f'''powershell -Command "$voice = New-Object -ComObject SAPI.SpVoice; $voice.Speak('Goodbye')"'''
            subprocess.run(ps_command, shell=True, capture_output=True, timeout=3)
        except:
            pass
    except Exception as e:
        print(f"\nError: {e}")
        try:
            text_clean = "Sorry, an error occurred".replace('"', '\\"').replace("'", "''")
            ps_command = f'''powershell -Command "$voice = New-Object -ComObject SAPI.SpVoice; $voice.Speak('Sorry, an error occurred')"'''
            subprocess.run(ps_command, shell=True, capture_output=True, timeout=3)
        except:
            pass