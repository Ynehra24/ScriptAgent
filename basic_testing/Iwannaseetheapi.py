from google import genai
from google.genai import types
from dotenv import load_dotenv
import os

load_dotenv()
client = genai.Client(api_key=os.environ["GEMINI_KEY"])

for model in [
    "gemini-2.5-flash-lite",
    "gemini-2.0-flash",
    "gemini-2.0-flash-lite",
    "gemini-2.5-pro",
    "gemini-2.5-flash",
]:
    try:
        r = client.models.generate_content(
            model=model,
            contents="Return OK.",
            config=types.GenerateContentConfig(max_output_tokens=10),
        )
        print("OK:", model, r.text)
    except Exception as e:
        print("FAIL:", model, e)