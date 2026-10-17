import pytest
from types import SimpleNamespace

from sokolov.openai_assets import is_obvious_case
from sokolov.openrouter_assets import openrouter_request

def test_is_obvious_case():
    # Test 1: Expression "all of them"
    body = "We invited ten friends, and all of them came."
    span = (35, 39)  # The span of 'them'
    assert is_obvious_case(body, span) == True

    # Test 2: No special expression
    body = "They don't like this."
    span = (0, 4)
    assert is_obvious_case(body, span) == False

    # Test 3: Expression "they were all"
    body = "They were all excited about the trip."
    span = (0, 4)
    assert is_obvious_case(body, span) == True

    # Test 4: Expression "they are all"
    body = "They are all coming to the party."
    span = (0, 4)
    assert is_obvious_case(body, span) == True

    # Test 5: Expression "they will all"
    body = "They will all be there at the meeting."
    span = (0, 4)
    assert is_obvious_case(body, span) == True

    # Test 6: Expression "some of them"
    body = "Some of them were late to the event."
    span = (0, 4)
    assert is_obvious_case(body, span) == True

    # Test 7: No special expression, similar structure
    body = "Most will arrive by noon."
    span = (0, 4)
    assert is_obvious_case(body, span) == False

    body = "There are more than just three of them, you know"
    span= (34, 38)
    assert is_obvious_case(body, span) == True

    body = "They resemble each other"
    span = (0, 4)
    assert is_obvious_case(body, span) == True

    body = "Last time I check they loved one another."
    span = (18, 22)
    assert is_obvious_case(body, span) == True

    body = '''I think  they have a new site up, they all should just go there and be done with it.'''
    span = (34, 38)
    assert is_obvious_case(body, span) == True

    body = "All of them!"
    span = (7, 11)
    assert is_obvious_case(body, span) == True


def fake_response(body, status_code=200):
    import json
    return SimpleNamespace(
        status_code=status_code,
        iter_content=lambda chunk_size=None: iter([json.dumps(body).encode()]),
        close=lambda: None,
    )


def test_openrouter_request_enables_web_search_without_system_message(monkeypatch):
    monkeypatch.setenv("OPENROUTER_API_KEY", "test-key")
    captured = {}
    response_body = {"choices": [{"message": {"content": "plural"}}]}

    def fake_post(url, **kwargs):
        captured["url"] = url
        captured.update(kwargs)
        return fake_response(response_body)

    monkeypatch.setattr("sokolov.openrouter_assets.requests.post", fake_post)

    result = openrouter_request(
        prompt="Classify this example.",
        system_message="",
        model="xiaomi/mimo-v2.6-pro",
        web_search_required=True,
    )

    assert result == "plural"
    assert captured["json"]["messages"] == [
        {"role": "user", "content": "Classify this example."}
    ]
    assert captured["json"]["tools"] == [{"type": "openrouter:web_search"}]

def test_openrouter_request_raises_on_null_content(monkeypatch):
    monkeypatch.setenv("OPENROUTER_API_KEY", "test-key")
    response_body = {"choices": [{"message": {"content": None}}]}
    captured = {}

    def fake_post(url, **kwargs):
        captured.update(kwargs)
        return fake_response(response_body)

    monkeypatch.setattr("sokolov.openrouter_assets.requests.post", fake_post)

    with pytest.raises(Exception, match="no message content"):
        openrouter_request(prompt="x", system_message="", model="moonshotai/kimi-k3")
    assert captured["timeout"] == (30, 600)
    assert captured["stream"] is True


def test_permalink_run_resume_fills_only_missing_rows(monkeypatch, tmp_path):
    import argparse
    import sys
    from pathlib import Path

    import pandas as pd

    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "sokolov"))
    import shagohod

    prompt = tmp_path / "prompt.txt"
    prompt.write_text("Classify {{TEXT}} at {{URL}}", encoding="utf-8")
    args = argparse.Namespace(
        promptstrat="context-via-permalink", llm="moonshotai/kimi-k3",
        promptfile=str(prompt), limit=None, outputdir=str(tmp_path), resume=True,
    )
    td = pd.DataFrame({
        "comment_body": ["they went", "they said", "they ran", "they sat"],
        "span": ["(0, 4)"] * 4,
        "permalink": ["https://example.com/x"] * 4,
    })

    # partial run: row 1 got a null-content reply, rows 2-3 were never reached
    partial = td.copy()
    partial["LLM_response"] = [
        "kept answer\nsingular", "None", "The request was rejected because it was considered high risk", "",
    ]
    partial["LLM_annotation"] = ["singular_they", "unknown_they", "", ""]
    partial.to_csv(shagohod.get_output_path(args, 1, in_progress=True), index=False)

    asked = []

    def fake_ask(**kwargs):
        asked.append(kwargs["prompt_text"])
        return "new answer\nplural", None

    monkeypatch.setattr(shagohod, "get_client", lambda model: None)
    monkeypatch.setattr(shagohod, "ask_llm_text", fake_ask)

    shagohod.run_context_permalink_zero_shot(td, args, run=1)

    assert len(asked) == 3
    assert "%they% said" in asked[0]
    assert "%they% ran" in asked[1]
    final_p = shagohod.get_output_path(args, 1, in_progress=False)
    result = pd.read_csv(final_p, keep_default_na=False)
    assert result["LLM_response"].tolist() == ["kept answer\nsingular"] + ["new answer\nplural"] * 3
    assert not Path(shagohod.get_output_path(args, 1, in_progress=True)).exists()

    # resuming a complete run makes no requests
    asked.clear()
    shagohod.run_context_permalink_zero_shot(td, args, run=1)
    assert asked == []


def test_openrouter_request_enforces_total_deadline(monkeypatch):
    monkeypatch.setenv("OPENROUTER_API_KEY", "test-key")
    clock = iter([0.0, 5.0, 11.0])
    monkeypatch.setattr("sokolov.openrouter_assets.time.monotonic", lambda: next(clock))

    def fake_post(url, **kwargs):
        # keep-alive whitespace trickling in past the deadline
        return SimpleNamespace(status_code=200, iter_content=lambda chunk_size=None: iter([b" ", b" "]), close=lambda: None)

    monkeypatch.setattr("sokolov.openrouter_assets.requests.post", fake_post)

    with pytest.raises(TimeoutError):
        openrouter_request(prompt="x", system_message="", model="xiaomi/mimo-v2.6-pro", timeout=10)
