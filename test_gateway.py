
import os
import json
import urllib.request
import urllib.error
from dotenv import load_dotenv

load_dotenv()

url = ""
key = os.getenv("LLM_API_KEY")

payload = {
    "model": "gpt-6-luna",
    "input": "I was billed twice.",
    "questions": [
        {
            "name": "team",
            "type": "choice",
            "choices": [
                {"value": "billing"},
                {"value": "technical"},
                {"value": "general"}
            ]
        }
    ]
}

request = urllib.request.Request(
    url,
    data=json.dumps(payload).encode("utf-8"),
    headers={
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json"
    },
    method="POST"
)

try:
    with urllib.request.urlopen(request, timeout=30) as response:
        print("Status:", response.status)
        print(response.read().decode())

except urllib.error.HTTPError as e:
    print("HTTP Status:", e.code)
    print("Response:", e.read().decode())

except Exception as e:
    print("Error:", str(e))
