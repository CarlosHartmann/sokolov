'''
August 2025: About to fly to ISLE in Santiago to present the first serious run of human vs machine at this task.
If at least somewhat successful, this will be the Shagohod; and Metal Gear will rise thereafter.
Cleaned in October 2025 to ensure all code is understood and (hopefully) commented properly for posterity.
'''

import re
import os
import ast
import time
import random
import argparse
import pandas as pd

from pathlib import Path

# to check provided arguments in the execution command
from sokolov.argparse_assets import dir_path

from typing import Optional, Tuple, Dict, Any

# AI company APIs
from openai import OpenAI
from anthropic import Anthropic 

LABELS = {"plural", "generic", "singular"}
LABEL_REGEX = re.compile(r"\b(plural|generic|singular)\b", re.IGNORECASE)
MORE_CONTEXT_REGEX = re.compile(r"\bmore\s+context\b[.!?)\"'\]]*\s*$", re.IGNORECASE)


def define_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()

    parser.add_argument('--promptfile', '-PF', required=True,
                        help="The file containing the prompt design to be used.")
    parser.add_argument('--promptstrat', '-PS', required=True,
                        help="The prompting strategy to be used. This should coincide with a function here defined that can work with a given promptfile and LLM.")
    parser.add_argument('--llm', '-LLM', type=str, required=True,
                        help="The LLM to be used for the experiment.")
    parser.add_argument('--runs', "-R", type=int, required=True,
                        help="Number of runs to be performed")
    parser.add_argument('--datapath', '-DP', type=dir_path, required=True,
                        help="Supply a working path to a directlry that contains the folders 'context', 'results', and the data themself named 'data_w_context.xlsw'.")
    
    parser.add_argument('--limit', '-L', type=int, required=False, default=None,
                        help="Limit the number of rows to be processed.")
    
    # currently not implemented. Shagohod is intended as a limited-run script using data that are known to be moderation-safe.
    #parser.add_argument('--moderation', '-M', action="store_true",
    #                    help="Send all comments to openAI moderation to check if they are acceptable to be used on GPT.")

    return parser


def handle_args() -> argparse.Namespace:
    """Handle argument-related edge cases by throwing meaningful errors."""
    parser = define_parser()
    args = parser.parse_args()

    args.testdata_file = os.path.join(args.datapath, "data_w_context.xlsx")
    args.context_path = os.path.join(args.datapath, "context")
    args.output_path = os.path.join(args.datapath, "results")

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
    
    if args.llm.startswith('claude'):
        if not os.getenv('ANTHROPIC_API_KEY'):
            print("ANTHROPIC_API_KEY environment variable required for Claude models")
            exit()
    else:
        if not os.getenv('OPENAI_API_KEY'):
            print("OPENAI_API_KEY environment variable required for OpenAI models")
            exit()
    
    return args


def parse_span(span_val: str) -> Tuple[int, int]:
    """
    Accepts '[start, end]' strings, lists/tuples, or pandas-sourced strings with spaces.
    Raises ValueError on malformed input.
    """

    if isinstance(span_val, str):
        try:
            # safer than manual split; tolerates spaces: "[ 12, 34 ]"
            start, end = ast.literal_eval(span_val)
        except Exception:
            # fallback: strip and split
            start, end = map(int, span_val.strip("[]").split(","))
    else:
        raise ValueError(f"Unsupported span type: {type(span_val)}")

    # convert to int (should work if provided span is valid)
    start, end = int(start), int(end)

    # Ensure the given span is indeed beginning:end
    if start >= end:
        raise ValueError(f"span start >= end: {start} >= {end}")
    return start, end


def mark_span(text: str, start: int, end: int, marker: str = "%") -> str:
    '''
    Marks an indicated part of a string.
    This will be used to highlight the pronoun to be disambiguated by the LLM.
    Example:
    mark_span("hello world", 6, 11)  # → "hello %world%"
    '''
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
    client,  # Accepts either OpenAI or Anthropic client
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
    Send prompt_text to either OpenAI or Anthropic with robust extraction and retries.
    """
    extra = extra or {}
    last_exc: Optional[Exception] = None
    is_claude = model.startswith('claude')

    for attempt in range(retries):
        try:
            if is_claude:
                # Anthropic API call
                kwargs = {
                    "model": model,
                    "max_tokens": max_output_tokens or 1024,
                    "messages": [{"role": "user", "content": prompt_text}],
                    **extra
                }
                if temperature is not None:
                    kwargs["temperature"] = temperature
                
                resp = client.messages.create(**kwargs)
                # Extract text from Anthropic response
                text = resp.content[0].text if resp.content else None
                
            else:
                # OpenAI API call (your existing logic)
                if hasattr(client, "responses"):  # modern SDK
                    kwargs = dict(model=model, input=prompt_text, **extra)
                    if max_output_tokens is not None:
                        kwargs["max_output_tokens"] = max_output_tokens
                    if temperature is not None:
                        kwargs["temperature"] = temperature
                    if seed is not None:
                        kwargs["seed"] = seed
                    resp = client.responses.create(**kwargs)
                else:  # fallback to Chat Completions
                    kwargs = dict(
                        model=model,
                        messages=[{"role": "user", "content": prompt_text}],
                        **extra,
                    )
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
            sleep_s = (base_sleep * (2 ** attempt)) * (1.0 + random.uniform(-jitter, jitter))
            time.sleep(max(0.0, sleep_s))

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


def _replace_first(s: str, old: str, new: str) -> str:
    """Replace only the first occurrence of `old` in `s`."""
    return s.replace(old, new, 1)


def _fill_prompt_with_sentence_and_url(prompt_template: str, sentence: str, url: Optional[str]) -> str:
    """
    Fill a template that may have:
      - {{TEXT}} for the sentence
      - {{URL}} / {{LINK}} / {{PERMALINK}} for the URL
    Fallbacks:
      - If there are two {{TEXT}} (or two "{}"), first -> sentence, second -> URL
      - If URL placeholder missing, append a permalink line at the end
    """
    tpl = prompt_template
    url = "" if url is None or (isinstance(url, float) and pd.isna(url)) else str(url)

    # Case 1: explicit URL placeholder present
    url_placeholders = ("{{URL}}", "{{LINK}}", "{{PERMALINK}}")
    if any(ph in tpl for ph in url_placeholders):
        for ph in url_placeholders:
            if ph in tpl:
                tpl = tpl.replace(ph, url)
        # now fill sentence placeholder(s)
        if "{{TEXT}}" in tpl:
            tpl = tpl.replace("{{TEXT}}", sentence)
        elif "{}" in tpl:
            tpl = tpl.replace("{}", sentence, 1)
        else:
            tpl = f"{tpl}\n\nText in question:\n{sentence}"
        return tpl

    # Case 2: no explicit URL placeholder — try a two-{{TEXT}} (or two-{}) pattern
    if "{{TEXT}}" in tpl:
        count = tpl.count("{{TEXT}}")
        if count >= 2:
            tpl = _replace_first(tpl, "{{TEXT}}", sentence)
            tpl = _replace_first(tpl, "{{TEXT}}", url)
            return tpl
        else:
            tpl = tpl.replace("{{TEXT}}", sentence)
            if url:
                tpl += f"\n\nPermalink for context: {url}"
            return tpl

    if "{}" in tpl:
        count = tpl.count("{}")
        if count >= 2:
            tpl = _replace_first(tpl, "{}", sentence)
            tpl = _replace_first(tpl, "{}", url)
            return tpl
        else:
            tpl = tpl.replace("{}", sentence, 1)
            if url:
                tpl += f"\n\nPermalink for context: {url}"
            return tpl

    # Case 3: no recognized placeholders at all — append both pieces
    extra = f"\n\nText in question:\n{sentence}"
    if url:
        extra += f"\n\nPermalink for context: {url}"
    return tpl + extra


def get_client(model_name: str):
    """Return appropriate client based on model name"""
    if model_name.startswith('claude'):
        return Anthropic()  # Uses ANTHROPIC_API_KEY env var
    else:
        return OpenAI()  # Uses OPENAI_API_KEY env var


def run_context_agnostic_zero_shot(td: pd.DataFrame, args: argparse.Namespace, run: int) -> Path:
    """
    Run a context-agnostic zero-shot experiment. Uses `ask_llm_text` for prompt sending.
    """
    client = get_client(args.llm)

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
            start, end = parse_span(span)
            they_sentence = mark_span(text, start, end, marker="%")
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
    base = f"results_{args.promptstrat}_{args.llm}_run{run}"
    if getattr(args, "limit", None):
        base += f"_top{args.limit}"
    output_filepath = os.path.join(args.output_path, f"{base}.xlsx")
    td.to_excel(output_filepath, index=False)

    if errors:
        print(f"Completed with {len(errors)} span/parsing errors (saved in sheet).")
    print(f"Results saved to {output_filepath}")


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
    client = get_client(args.llm)
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

    # Iterate rows
    errors = []
    for idx, text, span, item_id in td[["comment_body", "span", id_col]].itertuples(index=True, name=None):
        # Build marked sentence
        try:
            start, end = parse_span(span)
            they_sentence = mark_span(text, start, end, marker="%")
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
                ctx_text, ctx_path = _load_context(args.context_path, item_id)
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
    base = f"results_{args.promptstrat}_{args.llm}_run{run}"
    if getattr(args, "limit", None):
        base += f"_top{args.limit}"
    output_p = os.path.join(args.output_path, f"{base}.xlsx")
    td.to_excel(output_p, index=False)

    if errors:
        print(f"Completed with {len(errors)} span/context errors (saved in sheet).")
    print(f"Results saved to {output_p}")


def run_context_permalink_zero_shot(td: pd.DataFrame, args: argparse.Namespace, run: int) -> Path:
    """
    Like run_context_agnostic_zero_shot, but also injects a permalink from td['permalink']
    into the prompt template (supports {{URL}}/{{LINK}}/{{PERMALINK}} or a second {{TEXT}}/{}).
    """
    client = get_client(args.llm)

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

    # Prompt template
    prompt_template = Path(args.promptfile).read_text(encoding="utf-8")

    errors = []

    # Iterate rows (requires 'comment_body', 'span', 'permalink')
    cols_needed = ["comment_body", "span", "permalink"]
    missing = [c for c in cols_needed if c not in td.columns]
    if missing:
        raise KeyError(f"Missing required columns for permalink runner: {missing}")

    for idx, text, span, permalink in td[cols_needed].itertuples(index=True, name=None):
        try:
            start, end = parse_span(span)
            they_sentence = mark_span(text, start, end, marker="%")
        except Exception as e:
            td.at[idx, "LLM_response"] = f"[span_error] {e}"
            td.at[idx, "LLM_annotation"] = "unknown_they"
            errors.append((idx, str(e)))
            continue

        # Fill both placeholders (sentence + permalink)
        prompt_filled = _fill_prompt_with_sentence_and_url(prompt_template, they_sentence, permalink)

        # Send to model (reuses your robust helper)
        response_text, err = ask_llm_text(
            client=client,
            model=args.llm,
            prompt_text=prompt_filled,
            retries=getattr(args, "retries", 3),
            base_sleep=getattr(args, "base_sleep", 1.0),
            # If you want “no limit”, pass None (or leave commented)
            # max_output_tokens=getattr(args, "max_output_tokens", None),
            seed=getattr(args, "seed", None),
            extra=getattr(args, "openai_extra", None) or {},
        )

        # Guardrail against object reprs / empties
        if not response_text or response_text.startswith("Response("):
            response_text = "[no_text_returned]"

        td.at[idx, "LLM_response"] = response_text
        td.at[idx, "LLM_annotation"] = _extract_label(response_text)

    # Output (keeps your filename pattern)
    base = f"results_{args.promptstrat}_{args.llm}_run{run}"
    if getattr(args, "limit", None):
        base += f"_top{args.limit}"
    output_filepath = os.path.join(args.output_path, f"{base}.xlsx")
    td.to_excel(output_filepath, index=False)

    if errors:
        print(f"Completed with {len(errors)} span/parsing errors (saved in sheet).")
    print(f"Results saved to {output_filepath}")

def main():
    args = handle_args()

    # load the test data, the top row in the file are headers for the dataframe 
    td = pd.read_excel(args.testdata_file, sheet_name='data', engine='openpyxl')
    
    if args.limit is not None:
        td = td.head(args.limit)

    if args.promptstrat == "context-agnostic_zero-shot":
        for num in range(args.runs):
            run_context_agnostic_zero_shot(td, args, run=num+1)
    elif args.promptstrat == "context-ondemand_zero-shot":
        for num in range(args.runs):
            run_context_ondemand_zero_shot(td, args, run=num+1)
    elif args.promptstrat == "context-via-permalink":
        for num in range(args.runs):
            run_context_permalink_zero_shot(td, args, run=num+1)
    else:
        print(f"Unknown prompting strategy: {args.promptstrat}")
        exit()

if __name__ == "__main__":
    main()