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

MORE_CONTEXT_REGEX = re.compile(r"\bmore\s+context\b[.!?)\"'\]]*\s*$", re.IGNORECASE)

import getpass
username = getpass.getuser()
gdrive_user = 'GoogleDrive-carlos.hartm@gmail.com/Meine Ablage' if username == 'Carlitos' else 'GoogleDrive-smogshaik.uni@gmail.com/.shortcut-targets-by-id/1xXrtqarel363zcD11O0OzfvTnzROT6mb'


testdata_path = f'/Users/{username}/Library/CloudStorage/{gdrive_user}/2 - Uni-Ablage/04 sg.they/05 – Studies/2 - THEY Disambiguation/6 - study proper/LLMs'
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
    parser.add_argument('--runs', "-R", type=int, required=True,
                        help="Number of runs to be performed")
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
    
    if not args.runs:
        print("Specify number of runs")
        exit()
    elif args.runs == 0:
        print("Ha ha. Very funny.")
        exit()
    elif args.runs > 15:
        print("This many runs is not wise. Limit is 15.")
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
    max_output_tokens: int | None = None,
    allow_reasoning: bool = False,
    chat_fallback_model: str = "gpt-4o-mini",
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


def _needs_more_context(response_text: Optional[str]) -> bool:
    if not response_text or not isinstance(response_text, str):
        return False
    # True if the last meaningful words are "more context" (punctuation-insensitive)
    return bool(MORE_CONTEXT_REGEX.search(response_text.strip()))


def _load_context(context_dir: Path, id_value: Any) -> Tuple[str, Path]:
    """
    Load context text for this item. Tries exact filename match and `<ID>.txt`.
    Returns (context_text, path). Raises FileNotFoundError if missing.
    """
    candidates = [os.path.join(context_dir, str(id_value)), os.path.join(context_dir, f"{id_value}.txt)")]
    for p in candidates:
        if os.path.isfile(p):
            with open(p, 'r', encoding='utf-8') as file:
                text = file.read()
            return text, p
    raise FileNotFoundError(f"No context file for ID {id_value} in {context_dir}")


def run_context_agnostic_zero_shot(td: pd.DataFrame, args: argparse.Namespace, run: int) -> Path:
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
            #max_output_tokens=getattr(args, "max_output_tokens", 1024),
            #temperature=getattr(args, "temperature", None),
            seed=getattr(args, "seed", None),
            extra=getattr(args, "openai_extra", None) or {},
        )

        td.at[idx, "LLM_response"] = response_text
        td.at[idx, "LLM_annotation"] = _extract_label(response_text)

    # Output
    global output_path
    base = f"results_{args.promptstrat}_{args.llm}_run{run}"
    if getattr(args, "limit", None):
        base += f"_top{args.limit}"
    output_filepath = os.path.join(output_path, f"{base}.xlsx")
    td.to_excel(output_filepath, index=False)

    if errors:
        print(f"Completed with {len(errors)} span/parsing errors (saved in sheet).")
    print(f"Results saved to {output_path}")
    return output_path


def run_context_ondemand_zero_shot(td: pd.DataFrame, args: argparse.Namespace, run: int) -> Path:
    """
    Zero-shot with context-on-demand:
      1) Ask with base prompt. If the reply ends with 'more context', load the per-ID context
         and ask again (same sentence + appended context) with a strict 'return one word' instruction.
      2) Final response must end with plural|generic|singular; we then annotate like before.
    Expects args:
      - promptfile: path to the base prompt template
      - llm: model id
      - promptstrat: label for naming outputs
      - context_dir: directory with files named exactly like ID values (optionally with .txt)
      - (optional) id_column: name of the ID column (default "ID", fallback "id")
      - (optional) limit, output_dir, retries, base_sleep, max_output_tokens, temperature, seed, chat_fallback_model
    """
    client = OpenAI()
    td = td.copy()

    # Respect --limit
    if getattr(args, "limit", None):
        td = td.head(args.limit).copy()

    # Columns & dtypes (avoid FutureWarning)
    for col, dtype in [
        ("LLM_response", "string"),
        ("LLM_annotation", "string"),
        ("LLM_first_response", "string"),
    ]:
        if col not in td.columns:
            td[col] = pd.Series(pd.NA, index=td.index, dtype=dtype)
        else:
            td[col] = td[col].astype(dtype)

    # Track whether context was used
    if "LLM_more_context" not in td.columns:
        td["LLM_more_context"] = pd.Series(pd.NA, index=td.index, dtype="boolean")
    else:
        # cast to pandas nullable boolean (True/False/<NA>)
        td["LLM_more_context"] = td["LLM_more_context"].astype("boolean")

    # Prompt template & placeholder
    prompt_template = Path(args.promptfile).read_text(encoding="utf-8")
    placeholder = "{{TEXT}}" if "{{TEXT}}" in prompt_template else "{}"

    # ID column & context directory
    id_col = getattr(args, "id_column", None) or ("ID" if "ID" in td.columns else "id")
    if id_col not in td.columns:
        raise KeyError(f"ID column '{id_col}' not found in dataframe.")
    global context_path

    # Iterate rows
    errors = []
    for idx, text, span, item_id in td[["comment_body", "span", id_col]].itertuples(index=True, name=None):
        # Build marked sentence
        try:
            start, end = _parse_span(span)
            they_sentence = _mark_span(text, start, end, marker="%")
        except Exception as e:
            td.at[idx, "LLM_first_response"] = f"[span_error] {e}"
            td.at[idx, "LLM_response"] = f"[span_error] {e}"
            td.at[idx, "LLM_annotation"] = "unknown_they"
            td.at[idx, "LLM_more_context"] = False
            errors.append((idx, str(e)))
            continue

        # First turn
        prompt1 = prompt_template.replace(placeholder, they_sentence)
        resp1, err1 = ask_llm_text(
            client=client,
            model=args.llm,
            prompt_text=prompt1,
            retries=getattr(args, "retries", 3),
            base_sleep=getattr(args, "base_sleep", 1.0),
            #max_output_tokens=getattr(args, "max_output_tokens", 1024),   # give headroom
            #temperature=getattr(args, "temperature", 0),
            allow_reasoning=True,
            chat_fallback_model=getattr(args, "chat_fallback_model", "gpt-4o-mini"),
        )
        td.at[idx, "LLM_first_response"] = resp1

        # If more context requested, load and ask again
        if _needs_more_context(resp1):
            try:
                ctx_text, ctx_path = _load_context(context_path, item_id)
            except FileNotFoundError as e:
                td.at[idx, "LLM_response"] = f"[context_missing] {e}"
                td.at[idx, "LLM_annotation"] = "unknown_they"
                td.at[idx, "LLM_more_context"] = True
                errors.append((idx, str(e)))
                continue

            # Follow-up prompt: keep the same sentence, append context, and force a single-word decision
            prompt2 = (
                f"{prompt1}\n\n"
                f"The previous model requested more context so here it is below. Note that now you may no longer request more context as there is none that could be provided.\n"
                f"---\n{ctx_text}\n---\n\n"
                f"Given this context, identify which of the three categories the 'they' in question belongs to and have your answer end with the respective word: plural, generic, or singular."
            )

            resp2, err2 = ask_llm_text(
                client=client,
                model=args.llm,
                prompt_text=prompt2,
                retries=max(1, getattr(args, "retries", 3)),  # one quick retry is enough here
                base_sleep=getattr(args, "base_sleep", 1.0),
                #max_output_tokens=max(512, getattr(args, "max_output_tokens", 1024)),
                #temperature=getattr(args, "temperature", 0),
                allow_reasoning=True,
                chat_fallback_model=getattr(args, "chat_fallback_model", "gpt-4o-mini"),
            )

            # Use the second response as the official one
            final_text = resp2
            td.at[idx, "LLM_more_context"] = True
        else:
            final_text = resp1
            td.at[idx, "LLM_more_context"] = False

        # Guardrail: if something weird slips through (object repr, empty), make it obvious
        if not final_text or final_text.startswith("Response("):
            final_text = "[no_text_returned]"

        td.at[idx, "LLM_response"] = final_text
        td.at[idx, "LLM_annotation"] = _extract_label(final_text)

    # Write output
    global output_path
    base = f"results_{args.promptstrat}_{args.llm}_run{run}"
    if getattr(args, "limit", None):
        base += f"_top{args.limit}"
    output_path = os.path.join(output_path, f"{base}.xlsx")
    td.to_excel(output_path, index=False)

    if errors:
        print(f"Completed with {len(errors)} span/context errors (saved in sheet).")
    print(f"Results saved to {output_path}")
    return output_path


def main():
    global td
    args = handle_args()
    if args.limit is not None:
        td = td.head(args.limit)

    if args.promptstrat == "context-agnostic_zero-shot":
        for num in range(args.runs):
            run_context_agnostic_zero_shot(td, args, run=num+1)
    elif args.promptstrat == "context-ondemand_zero-shot":
        for num in range(args.runs):
            run_context_ondemand_zero_shot(td, args, run=num+1)
    else:
        print(f"Unknown prompting strategy: {args.promptstrat}")
        exit(1)

if __name__ == "__main__":
    main()