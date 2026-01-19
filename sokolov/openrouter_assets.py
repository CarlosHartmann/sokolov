'''
sokolov.openrouter_assets: Functions and types that I need for my OpenRouter implementation.
'''

import requests
import json
import os

# load API key from environment variable
OPENROUTER_API_KEY = os.environ.get('OPENROUTER_API_KEY')

def openrouter_request(prompt: str, system_message: str, model: str) -> str:

    if system_message == "":
        raise ValueError("System message cannot be empty for OpenRouter requests.")

    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
    "model": model,
    "messages": [
    {
    "role": "system",
    "content": system_message
    },
    {
    "role": "user",
    "content": prompt
    }
    ],
    "temperature": 0
    }

    response = requests.post(url, headers=headers, json=payload)
    response_json = response.json()
    return response_json['choices'][0]['message']['content']