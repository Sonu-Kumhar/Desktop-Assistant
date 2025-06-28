import os, requests
from dotenv import load_dotenv

load_dotenv()                                   # loads .env
RAPID_KEY = os.getenv("RAPIDAPI_KEY")
if not RAPID_KEY:
    raise EnvironmentError("RAPIDAPI_KEY missing in .env")

URL = "https://chatgpt-42.p.rapidapi.com/chat"
HEADERS = {
    "Content-Type": "application/json",
    "x-rapidapi-key": RAPID_KEY,
    "x-rapidapi-host": "chatgpt-42.p.rapidapi.com",
}
MODEL = "gpt-4o-mini"

def get_reply(history: list[dict], user_msg: str, max_tokens: int = 150) -> str:
    """Return chatbot reply text."""
    payload = {
        "model": MODEL,
        "messages": history[-10:] + [{"role": "user", "content": user_msg}],
        "max_tokens": max_tokens,
        "temperature": 0.8,
    }
    r = requests.post(URL, json=payload, headers=HEADERS, timeout=20)
    r.raise_for_status()
    return r.json()["choices"][0]["message"]["content"].strip()
