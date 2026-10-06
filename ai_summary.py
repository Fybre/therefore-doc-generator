"""
ai_summary.py — Generate an AI narrative summary of a Therefore configuration.

Uses an OpenAI-compatible API (e.g. LM Studio running locally).
"""
from __future__ import annotations

# Reasoning models (OpenAI o-series, gpt-5 family) spend completion tokens on hidden
# reasoning before writing any text, so they need far more room than the visible output.
_REASONING_HEADROOM = 4000

# Request style that worked per (base_url, model), so repeated calls skip failed attempts.
_STYLE_CACHE: dict[tuple, str] = {}


def _chat(client, model: str, prompt: str, temperature: float, max_tokens: int) -> str:
    """
    Run one chat completion, adapting to what the endpoint accepts.

    Tries, in order:
      "classic"   — temperature + max_tokens (LM Studio, gpt-4o / gpt-4.1, most compatible servers)
      "reasoning" — max_completion_tokens with reasoning headroom, default temperature,
                    reasoning_effort="low" (OpenAI o-series / gpt-5)
      "plain"     — as "reasoning" but without reasoning_effort
    A style is only abandoned on a 400 error that names one of its parameters.
    """
    from openai import BadRequestError

    messages = [{"role": "user", "content": prompt}]
    styles = {
        "classic":   {"temperature": temperature, "max_tokens": max_tokens},
        "reasoning": {"max_completion_tokens": max_tokens + _REASONING_HEADROOM, "reasoning_effort": "low"},
        "plain":     {"max_completion_tokens": max_tokens + _REASONING_HEADROOM},
    }
    key = (str(client.base_url), model)
    order = list(styles)
    if key in _STYLE_CACHE:
        order.remove(_STYLE_CACHE[key])
        order.insert(0, _STYLE_CACHE[key])

    last_exc = None
    for style in order:
        try:
            response = client.chat.completions.create(model=model, messages=messages, **styles[style])
        except BadRequestError as exc:
            message = str(exc)
            if not any(param in message for param in styles[style]):
                raise
            last_exc = exc
            continue
        _STYLE_CACHE[key] = style
        text = (response.choices[0].message.content or "").strip()
        if not text:
            reason = response.choices[0].finish_reason
            raise RuntimeError(f"model returned no text (finish_reason={reason})")
        return text
    raise last_exc


def _build_prompt(categories, workflows, profiles, eforms, maps, server_info=None) -> str:
    from build_doc import get_name

    lines = [
        "You are writing the opening section of a technical design document for a Therefore "
        "document management system implementation. Your task is to write a concise executive "
        "summary — 2 to 3 short paragraphs — that a non-technical reader can use to understand "
        "what this system does and why it exists.",
        "",
        "Guidelines:",
        "- Infer the business purpose from the category and workflow names. What kinds of "
        "documents or processes does this system manage? What business domain does it serve?",
        "- Describe the overall shape of the system: what the main document types are, what "
        "the key workflows do, and how users interact with it (e.g. via eForms).",
        "- Do NOT list counts, field numbers, or task numbers. Do not enumerate every category "
        "or workflow by name unless it helps explain the purpose.",
        "- Do not use marketing language. Do not say 'streamline', 'robust', 'leverage', "
        "'comprehensive', 'enhance', or 'solution'.",
        "- Do not invent features not supported by the data below.",
        "- Write in plain prose. No bullet points, no headings, no bold text.",
        "",
        "SYSTEM CONFIGURATION",
        "====================",
    ]

    # Server context
    if server_info:
        if server_info.get("tenant_name"):
            lines.append(f"Tenant: {server_info['tenant_name']}")
        if server_info.get("region"):
            lines.append(f"Region: {server_info['region']}")

    # Categories — names only, no field counts
    cat_names = [name for _no, name, _folder, _cat in categories]
    lines.append(f"\nDocument categories ({len(cat_names)}): {', '.join(cat_names)}")

    # Workflows — names only
    wf_names = [get_name(wf) or "Unnamed" for wf in workflows]
    if wf_names:
        lines.append(f"\nWorkflow processes ({len(wf_names)}): {', '.join(wf_names)}")

    # eForms — names only
    ef_list = [e["name"] for e in (eforms or []) if e.get("name")]
    if ef_list:
        lines.append(f"\neForms: {', '.join(ef_list[:15])}")

    # Keyword dictionaries — names give domain context
    kw = maps.get("kw_dicts", [])
    if kw:
        kw_names = [d["name"] for d in kw if d.get("name")]
        lines.append(f"\nKeyword dictionaries: {', '.join(kw_names[:15])}")

    lines += ["", "Write the executive summary now:"]
    return "\n".join(lines)


def summarize_script(
    code: str,
    context: str = "",
    ai_url: str = "http://localhost:1234/v1",
    ai_model: str = None,
    api_key: str = "lm-studio",
) -> str:
    """
    Return a one-sentence plain-English description of what a script does.
    Raises on connection failure.
    """
    try:
        from openai import OpenAI
    except ImportError:
        raise RuntimeError("openai package not installed")

    client = OpenAI(base_url=ai_url, api_key=api_key or "lm-studio")
    prompt = (
        f"Describe in one concise sentence what the following Therefore script does. "
        f"Be specific about field names, conditions, or actions involved. "
        f"Do not use marketing language.\n"
        + (f"Context: {context}\n" if context else "")
        + f"\n{code}"
    )
    return _chat(client, ai_model or "local-model", prompt, temperature=0.2, max_tokens=120)


def generate_ai_summary(
    categories,
    workflows,
    profiles,
    eforms,
    maps,
    server_info=None,
    ai_url: str = "http://localhost:1234/v1",
    ai_model: str = None,
    api_key: str = "lm-studio",
    log_fn=None,
) -> str:
    """
    Call LM Studio (OpenAI-compatible) to generate a narrative system summary.
    Returns the summary text, or raises on failure.
    """
    try:
        from openai import OpenAI
    except ImportError:
        raise RuntimeError("openai package not installed — add it to requirements.txt")

    client = OpenAI(base_url=ai_url, api_key=api_key or "lm-studio")
    prompt = _build_prompt(categories, workflows, profiles, eforms, maps, server_info)

    if log_fn:
        log_fn(f"Requesting AI summary from {ai_url} ...")

    text = _chat(client, ai_model or "local-model", prompt, temperature=0.3, max_tokens=512)

    if log_fn:
        log_fn("AI summary received.")

    return text
