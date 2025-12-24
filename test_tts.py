# test_speech_now.py
import os
import time

print("="*60)
print("TESTING SPEECH - MUST WORK")
print("="*60)

# METHOD 1: PowerShell - MOST RELIABLE
print("\n1. Testing PowerShell speech...")
os.system('''powershell -Command "Add-Type -AssemblyName System.Speech; $speak = New-Object System.Speech.Synthesis.SpeechSynthesizer; $speak.Speak('Hello, this is a test')"''')
print("Did you hear 'Hello, this is a test'?")
time.sleep(2)

# METHOD 2: VBScript - ALWAYS WORKS ON WINDOWS
print("\n2. Testing VBScript speech...")
# Create a VBS file
vbs_code = '''
Set speech = CreateObject("SAPI.SpVoice")
speech.Speak "Hello from Windows speech"
'''

with open('speak_test.vbs', 'w') as f:
    f.write(vbs_code)

# Run it
os.system('cscript //nologo speak_test.vbs')

# Clean up
os.remove('speak_test.vbs')
print("Did you hear 'Hello from Windows speech'?")
time.sleep(2)

# METHOD 3: Direct command
print("\n3. Testing direct command...")
os.system('mshta vbscript:Execute("CreateObject(""SAPI.SpVoice"").Speak(""Testing three"")(window.close)")')
print("Did you hear 'Testing three'?")

print("\n" + "="*60)
print("RESULTS:")
print("If you heard ALL THREE, speech works!")
print("If you heard NONE, your speakers are OFF/muted.")
print("="*60)

# Test with simple input
print("\nNow testing interactive speech...")
time.sleep(1)

# Simple interactive test
import speech_recognition as sr

def simple_speak(text):
    """Simple speech that MUST work"""
    print(f"Speaking: {text}")
    # Use PowerShell
    text_clean = text.replace('"', "'")
    cmd = f'''powershell -Command "Add-Type -AssemblyName System.Speech; (New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak('{text_clean}')"'''
    os.system(cmd)

def listen():
    """Simple listen"""
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("\nSpeak something (say 'test' or 'hello')...")
        try:
            audio = r.listen(source, timeout=5)
            text = r.recognize_google(audio)
            return text.lower()
        except:
            return ""

# Test
simple_speak("Good afternoon! How are you?")
print("\nListening for your response...")
response = listen()

if response:
    print(f"You said: {response}")
    simple_speak(f"I heard you say {response}")
else:
    simple_speak("I didn't hear anything")

simple_speak("Speech test complete!")