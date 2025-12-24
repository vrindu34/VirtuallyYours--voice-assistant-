
# Virtually Yours – Python Voice-Based Virtual Assistant

Virtually Yours is a customizable, voice-controlled desktop virtual assistant built using Python.  
Users can personalize the assistant by assigning it a name of their choice, making the interaction more natural and user-centric.

The assistant enables hands-free system automation, information retrieval, calculations, weather and news updates, note management, and AI-powered question answering.

---

## Overview

Virtually Yours is designed to adapt to the user rather than enforce a fixed identity.  
During startup, users are prompted to customize the assistant’s name, after which all interactions are carried out using the chosen name.

The project demonstrates practical integration of speech recognition, text-to-speech synthesis, external APIs, and Windows system-level automation.

---

## Features

- Customizable assistant name defined by the user at runtime  
- Voice-based command recognition using SpeechRecognition  
- Text-to-speech responses using Windows SAPI  
- Wikipedia and Google search integration  
- AI-powered question answering using OpenAI  
- Mathematical and factual calculations using Wolfram Alpha  
- Real-time weather updates using OpenWeather API  
- Latest news headlines using News API  
- Voice-based note creation and retrieval  
- Countdown timer functionality  
- System-level controls (lock, shutdown, restart)  
- Opening websites and system applications  
- Location search using Google Maps  

---

## Tech Stack

**Language**
- Python  

**Libraries and APIs**
- SpeechRecognition  
- pyttsx3  
- OpenAI API  
- Wolfram Alpha API  
- OpenWeather API  
- News API  
- Requests  
- BeautifulSoup  
- WMI  
- python-dotenv  

---

## Platform Support

- Operating System: Windows  
- Uses Windows SAPI for text-to-speech  
- Uses Windows-specific system commands  

---

## Project Structure

```text
VirtuallyYours/
├── main.py                 # Main entry point for the virtual assistant
├── test_tts.py             # Text-to-speech testing module
├── assistant_notes.txt     # Stores user-created voice notes
├── requirements.txt        # Project dependencies
├── .gitignore              # Git ignore rules
└── README.md               # Project documentation
>>>>>>> 217fcd0 (Initial commit: Virtually Yours virtual assistant)
