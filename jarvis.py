import subprocess
import re
import json
import os
import time
import threading
import anthropic
import speech_recognition as sr
from pynput import keyboard
from dotenv import load_dotenv

load_dotenv()

client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
recognizer = sr.Recognizer()
STATE_FILE = os.path.expanduser("~/jarvis-mac/jarvis_state.json")
ELEVENLABS_API_KEY = os.getenv("ELEVENLABS_API_KEY", "")

def speak(text):
    lang = "th" if any('\u0e00' <= c <= '\u0e7f' for c in text) else "en"
    if lang == "en" and ELEVENLABS_API_KEY:
        try:
            from elevenlabs import ElevenLabs
            el = ElevenLabs(api_key=ELEVENLABS_API_KEY)
            audio = el.text_to_speech.convert(
                voice_id="pNInz6obpgDQGcFmaJgB",
                text=text,
                model_id="eleven_monolingual_v1",
                output_format="mp3_44100_128"
            )
            with open("/tmp/jarvis_speech.mp3", "wb") as f:
                for chunk in audio:
                    if chunk:
                        f.write(chunk)
            subprocess.run(["afplay", "/tmp/jarvis_speech.mp3"])
            return
        except Exception as e:
            print(f"ElevenLabs error: {e}")
    subprocess.run(["say", "-v", "Kanya", text])

def clean_text(text):
    text = re.sub(r'[^\u0000-\u007F\u0E00-\u0E7F\s]', '', text)
    text = re.sub(r'\*+', '', text)
    return text.strip()

def listen():
    try:
        with sr.Microphone() as source:
            print("กำลังฟัง... พูดได้เลย!")
            recognizer.adjust_for_ambient_noise(source, duration=0.3)
            audio = recognizer.listen(source, timeout=10, phrase_time_limit=8)
        text = recognizer.recognize_google(audio, language="th-TH")
        print(f"คุณ: {text}")
        return text
    except sr.WaitTimeoutError:
        print("หมดเวลา ลองใหม่")
        return None
    except sr.UnknownValueError:
        print("ไม่ได้ยิน ลองใหม่")
        return None
    except Exception as e:
        print(f"Error: {e}")
        return None

def ask_jarvis(question):
    print("Jarvis กำลังคิด...")
    keywords = ["ราคา", "หุ้น", "ข่าว", "วันนี้", "ล่าสุด", "ตอนนี้", "2025", "2026", "price", "news", "today", "latest", "current"]
    need_search = any(word in question.lower() for word in keywords)
    params = {
        "model": "claude-sonnet-4-6",
        "max_tokens": 1024,
        "system": "คุณคือ Jarvis AI ส่วนตัว ถ้าถามเรื่องง่ายตอบสั้น 1-2 ประโยค ถ้าถามซับซ้อนค่อยอธิบายยาว ห้ามใช้ bullet point ห้ามใช้ * ตอบเป็นประโยคธรรมชาติเหมือนเพื่อนคุยกัน ถ้าถามภาษาไทยตอบไทย ถ้าถามอังกฤษตอบอังกฤษ",
        "messages": [{"role": "user", "content": question}]
    }
    if need_search:
        print("กำลังค้นข้อมูล...")
        params["tools"] = [{"type": "web_search_20250305", "name": "web_search"}]
    message = client.messages.create(**params)
    answer = ""
    for block in message.content:
        if hasattr(block, "text"):
            answer += block.text
    answer = clean_text(answer)
    print(f"Jarvis: {answer}")
    speak(answer)
    with open(STATE_FILE, 'w') as f:
        json.dump({"status": "responded", "text": answer}, f, ensure_ascii=False)

def watch_state():
    last_text = ""
    while True:
        try:
            with open(STATE_FILE, "r") as f:
                state = json.load(f)
            text = state.get("text", "")
            source = state.get("file", "")
            if text and text != last_text and source:
                last_text = text
                print(f"Jarvis (ไฟล์): กำลังอ่านผลวิเคราะห์...")
                speak(text)
                with open(STATE_FILE, 'w') as f:
                    json.dump({"status": "spoken", "text": text}, f, ensure_ascii=False)
        except:
            pass
        time.sleep(1)

t = threading.Thread(target=watch_state, daemon=True)
t.start()

def on_press(key):
    try:
        if key.char == 'q':
            question = listen()
            if question:
                ask_jarvis(question)
    except:
        pass

print("Jarvis พร้อมแล้ว! กด Q เพื่อพูด, กด Ctrl+C เพื่อออก")
with keyboard.Listener(on_press=on_press) as listener:
    listener.join()