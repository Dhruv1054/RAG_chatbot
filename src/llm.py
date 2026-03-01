import json
import requests
from src.config import LLM_MODEL, LLM_TEMPERATURE, LLM_TIMEOUT, OLLAMA_BASE_URL

_GENERATE_URL = f"{OLLAMA_BASE_URL}/api/generate"
_TAGS_URL     = f"{OLLAMA_BASE_URL}/api/tags"


def _build_payload(prompt, model=LLM_MODEL, temperature=LLM_TEMPERATURE, stream=True):
    return {
        "model":   model,
        "prompt":  prompt,
        "stream":  stream,
        "options": {"temperature": temperature, "num_predict": 1024},
    }


def stream_response(prompt, model=LLM_MODEL, temperature=LLM_TEMPERATURE):
    payload = _build_payload(prompt, model=model, temperature=temperature, stream=True)
    try:
        with requests.post(_GENERATE_URL, json=payload, stream=True, timeout=LLM_TIMEOUT) as resp:
            resp.raise_for_status()
            for raw_line in resp.iter_lines():
                if not raw_line:
                    continue
                try:
                    data = json.loads(raw_line)
                except json.JSONDecodeError:
                    continue
                token = data.get("response", "")
                if token:
                    yield token
                if data.get("done", False):
                    break
    except requests.exceptions.ConnectionError:
        raise RuntimeError(
            f"Cannot connect to Ollama at {OLLAMA_BASE_URL}. Run: ollama serve"
        )
    except requests.exceptions.Timeout:
        raise RuntimeError(f"Ollama timed out after {LLM_TIMEOUT}s. Try again.")


def generate(prompt, model=LLM_MODEL, temperature=LLM_TEMPERATURE):
    return "".join(stream_response(prompt, model=model, temperature=temperature))


def check_ollama_health():
    try:
        resp = requests.get(_TAGS_URL, timeout=5)
        resp.raise_for_status()
        models = [m["name"] for m in resp.json().get("models", [])]
        if not any(LLM_MODEL in m for m in models):
            return False, f"Model '{LLM_MODEL}' not found. Run: ollama pull {LLM_MODEL}"
        return True, "ok"
    except requests.exceptions.ConnectionError:
        return False, f"Ollama is not running at {OLLAMA_BASE_URL}."
    except Exception as exc:
        return False, str(exc)
