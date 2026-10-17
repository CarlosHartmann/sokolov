'''
sokolov.openrouter_assets: Functions and types that I need for my OpenRouter implementation.
'''

import requests
import json
import os
import time

def openrouter_request(
    prompt: str,
    system_message: str,
    model: str,
    web_search_required: bool = False,
    timeout: float = 600,
) -> str:

    api_key = os.environ.get('OPENROUTER_API_KEY')
    if not api_key:
        raise ValueError("OPENROUTER_API_KEY environment variable is not set.")

    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    messages = []
    if system_message:
        messages.append({"role": "system", "content": system_message})
    messages.append({"role": "user", "content": prompt})

    payload = {
        "model": model,
        "messages": messages,
        "temperature": 0,
    }

    if web_search_required:
        payload["tools"] = [{"type": "openrouter:web_search"}]

    if "terminus" in model:
        # add boolean extra_body={"reasoning": {"effort": "high"} for terminus models
        payload["extra_body"] = {"reasoning": {"effort": "high"}}

    # OpenRouter may send keep-alive whitespace while a request is processing, which resets
    # requests' per-read timeout, so enforce a total deadline while reading the body.
    deadline = time.monotonic() + timeout
    response = requests.post(url, headers=headers, json=payload, timeout=(30, timeout), stream=True)
    body = bytearray()
    try:
        for chunk in response.iter_content(chunk_size=None):
            body.extend(chunk)
            if time.monotonic() > deadline:
                raise TimeoutError(f"OpenRouter request exceeded {timeout} seconds")
    finally:
        response.close()
    response_json = json.loads(body)

    if response.status_code != 200:
        raise Exception(f"OpenRouter API request failed with status code {response.status_code}: {response_json}")
    
    content = response_json['choices'][0]['message']['content']
    if content is None:
        # e.g. the model stopped after a tool call; raise so the caller retries
        raise Exception(f"OpenRouter returned no message content: {response_json}")
    return content