'''
August 2025: About to fly to ISLE in Santiago to present the first serious run of human vs machine at this task.
If at least somewhat successful, this will be the Shagohod; and Metal Gear will rise thereafter.
'''

from __future__ import annotations

import argparse
import os
import time
import random
from typing import Optional, Tuple, Dict, Any
from pathlib import Path
import ast
import re
from typing import Tuple, Optional

import pandas as pd
from openai import OpenAI

LABELS = {"plural", "generic", "singular"}
LABEL_REGEX = re.compile(r"\b(plural|generic|singular)\b", re.IGNORECASE)


testdata_path = '/Users/Carlitos/Library/CloudStorage/GoogleDrive-carlos.hartm@gmail.com/Meine Ablage/2 - Uni-Ablage/04 sg.they/05 – Studies/2 - THEY Disambiguation/6 - study proper/LLMs'
testdata_file = os.path.join(testdata_path, "data_w_context.xlsx")
context_path = os.path.join(testdata_path, "context")
output_path = os.path.join(testdata_path, "results")

# load the test data, the top row in the file are headers for the dataframe 
td = pd.read_excel(testdata_file, sheet_name='data', engine='openpyxl')

def define_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()

    parser.add_argument('--limit', '-L', type=int, required=False, default=None,
                        help="Limit the number of rows to be processed.")
    parser.add_argument('--promptfile', '-PF', required=False,
                        help="The file containing the prompt design to be used.")
    parser.add_argument('--promptstrat', '-PS', required=False,
                        help="The prompting strategy to be used. This should coincide with a function here defined that can work with a given promptfile and LLM.")
    parser.add_argument('--llm', '-LLM', type=str, required=False,
                        help="The LLM to be used for the experiment.")
    parser.add_argument('--moderation', '-M', action="store_true",
                        help="Send all comments to openAI moderation to check if they are acceptable to be used on GPT.")
    return parser


def handle_args() -> argparse.Namespace:
    """Handle argument-related edge cases by throwing meaningful errors."""
    parser = define_parser()
    args = parser.parse_args()

    if not args.llm:
        print("LLM to be used not set.")
        exit()
    
    return args


def _parse_span(span_val) -> Tuple[int, int]:
    """
    Accepts '[start, end]' strings, lists/tuples, or pandas-sourced strings with spaces.
    Raises ValueError on malformed input.
    """
    if span_val is None or (isinstance(span_val, float) and pd.isna(span_val)):
        raise ValueError("span is NaN/None")

    if isinstance(span_val, (list, tuple)) and len(span_val) == 2:
        start, end = span_val
    elif isinstance(span_val, str):
        try:
            # safer than manual split; tolerates spaces: "[ 12, 34 ]"
            start, end = ast.literal_eval(span_val)
        except Exception:
            # fallback: strip and split
            start, end = map(int, span_val.strip("[]").split(","))
    else:
        raise ValueError(f"Unsupported span type: {type(span_val)}")

    start, end = int(start), int(end)
    if start >= end:
        raise ValueError(f"span start >= end: {start} >= {end}")
    return start, end


def _mark_span(text: str, start: int, end: int, marker: str = "%") -> str:
    if not isinstance(text, str):
        raise ValueError("text is not a string")
    if start < 0 or end > len(text):
        raise ValueError(f"span out of bounds for text length {len(text)}: ({start}, {end})")
    return text[:start] + marker + text[start:end] + marker + text[end:]


def _extract_label(response_text: Optional[str]) -> str:
    if not response_text or not isinstance(response_text, str):
        return "unknown_they"
    matches = LABEL_REGEX.findall(response_text)
    return (matches[-1].lower() + "_they") if matches else "unknown_they"


def _extract_openai_text(resp: Any) -> Optional[str]:
    """
    Works with Responses API objects and falls back to raw repr if needed.
    """
    # Newer SDK (Responses API)
    text = getattr(resp, "output_text", None)
    if text:
        return text

    # Defensive fallback paths (SDK variations)
    output = getattr(resp, "output", None) or getattr(resp, "outputs", None)
    if output and len(output):
        # responses: outputs[0].content[0].text
        first = output[0]
        content = getattr(first, "content", None)
        if content and len(content):
            maybe = getattr(content[0], "text", None)
            if maybe:
                return maybe

    # Chat Completions (older SDK)
    choices = getattr(resp, "choices", None)
    if choices and len(choices):
        msg = getattr(choices[0], "message", None)
        if msg and getattr(msg, "content", None):
            return msg.content

    # Last resort
    return None


def ask_llm_text(
    client: OpenAI,
    model: str,
    prompt_text: str,
    retries: int = 3,
    base_sleep: float = 1.0,
    jitter: float = 0.25,
    max_output_tokens: Optional[int] = 256,
    temperature: Optional[float] = None,
    seed: Optional[int] = None,
    extra: Optional[Dict[str, Any]] = None,
) -> Tuple[str, Optional[Exception]]:
    """
    Send `prompt_text` to OpenAI with robust extraction and retries.
    Returns (response_text, error). If error is not None, response_text is a formatted error string.
    Compatible with both `client.responses.create` and the older `client.chat.completions.create`.
    """
    extra = extra or {}
    last_exc: Optional[Exception] = None

    for attempt in range(retries):
        try:
            if hasattr(client, "responses"):  # modern SDK
                kwargs = dict(model=model, input=prompt_text, **extra)
                if max_output_tokens is not None:
                    kwargs["max_output_tokens"] = max_output_tokens
                if temperature is not None:
                    kwargs["temperature"] = temperature
                if seed is not None:
                    kwargs["seed"] = seed
                resp = client.responses.create(**kwargs)
            else:  # fallback to Chat Completions (older SDK)
                kwargs = dict(
                    model=model,
                    messages=[{"role": "user", "content": prompt_text}],
                    **extra,
                )
                # chat.completions uses `max_tokens` (not max_output_tokens)
                if max_output_tokens is not None:
                    kwargs["max_tokens"] = max_output_tokens
                if temperature is not None:
                    kwargs["temperature"] = temperature
                if seed is not None:
                    kwargs["seed"] = seed
                resp = client.chat.completions.create(**kwargs)

            text = _extract_openai_text(resp)
            if text is None:
                text = str(resp)
            return text, None

        except Exception as e:
            last_exc = e
            # exponential backoff with jitter
            sleep_s = (base_sleep * (2 ** attempt)) * (1.0 + random.uniform(-jitter, jitter))
            time.sleep(max(0.0, sleep_s))

    # failed after retries
    return f"[api_error] {repr(last_exc)}", last_exc


def run_context_agnostic_zero_shot(td: pd.DataFrame, args: argparse.Namespace) -> Path:
    """
    Run a context-agnostic zero-shot experiment. Uses `ask_llm_text` for prompt sending.
    """
    client = OpenAI()

    # Respect --limit for processing
    if getattr(args, "limit", None):
        td = td.head(args.limit).copy()
    else:
        td = td.copy()

    # Ensure text cols are strings (prevents dtype warnings)
    for col in ("LLM_response", "LLM_annotation"):
        if col not in td.columns:
            td[col] = pd.Series(pd.NA, index=td.index, dtype="string")
        else:
            td[col] = td[col].astype("string")

    # Prompt template + placeholder
    prompt_template = Path(args.promptfile).read_text(encoding="utf-8")
    placeholder = "{{TEXT}}" if "{{TEXT}}" in prompt_template else "{}"

    # Iterate
    errors = []
    for idx, text, span in td[["comment_body", "span"]].itertuples(index=True, name=None):
        try:
            start, end = _parse_span(span)
            they_sentence = _mark_span(text, start, end, marker="%")
        except Exception as e:
            td.at[idx, "LLM_response"] = f"[span_error] {e}"
            td.at[idx, "LLM_annotation"] = "unknown_they"
            errors.append((idx, str(e)))
            continue

        prompt_filled = prompt_template.replace(placeholder, they_sentence)

        # Call the reusable helper (handles Responses vs Chat, retries, etc.)
        response_text, err = ask_llm_text(
            client=client,
            model=args.llm,
            prompt_text=prompt_filled,
            retries=getattr(args, "retries", 3),
            base_sleep=getattr(args, "base_sleep", 1.0),
            max_output_tokens=getattr(args, "max_output_tokens", 1024),
            temperature=getattr(args, "temperature", None),
            seed=getattr(args, "seed", None),
            extra=getattr(args, "openai_extra", None) or {},
        )

        td.at[idx, "LLM_response"] = response_text
        td.at[idx, "LLM_annotation"] = _extract_label(response_text)

    # Output
    global output_path
    base = f"results_{args.promptstrat}_{args.llm}"
    if getattr(args, "limit", None):
        base += f"_top{args.limit}"
    output_path = os.path.join(output_path, f"{base}.xlsx")
    td.to_excel(output_path, index=False)

    if errors:
        print(f"Completed with {len(errors)} span/parsing errors (saved in sheet).")
    print(f"Results saved to {output_path}")
    return output_path


def main():
    global td
    args = handle_args()
    if args.limit is not None:
        td = td.head(args.limit)

    if args.promptstrat == "context-agnostic_zero-shot":
        td = run_context_agnostic_zero_shot(td, args)
    else:
        print(f"Unknown prompting strategy: {args.promptstrat}")
        exit(1)

if __name__ == "__main__":
    main()