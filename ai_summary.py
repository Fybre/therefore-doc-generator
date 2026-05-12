"""
ai_summary.py — Generate an AI narrative summary of a Therefore configuration.

Uses an OpenAI-compatible API (e.g. LM Studio running locally).
"""
from __future__ import annotations


def _build_prompt(categories, workflows, profiles, eforms, maps, server_info=None) -> str:
    lines = [
        "You are a technical writer producing a Therefore document management system "
        "design document. Based on the configuration data below, write a 3-5 paragraph "
        "executive summary in clear professional prose. Describe what the system does, "
        "its main document categories, key workflow automation, eForm usage, and any "
        "notable integrations or configuration details. Do not use bullet points or "
        "headings — flowing paragraphs only.",
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
        if server_info.get("server_name"):
            lines.append(f"Server: {server_info['server_name']}")
        if server_info.get("service_version"):
            lines.append(f"Service version: {server_info['service_version']}")

    # Categories
    lines.append(f"\nCategories ({len(categories)}):")
    for no, name, _folder, cat in categories:
        fields_el = cat.find("Fields")
        field_count = len(list(fields_el)) if fields_el is not None else 0
        lines.append(f"  - {name} ({field_count} fields)")

    # Workflows
    lines.append(f"\nWorkflow Processes ({len(workflows)}):")
    for wf in workflows:
        from build_doc import get_name
        name = get_name(wf) or "Unnamed"
        tasks_el = wf.find("Tasks")
        task_count = len(list(tasks_el)) if tasks_el is not None else 0
        lines.append(f"  - {name} ({task_count} tasks)")

    # Indexing profiles
    lines.append(f"\nIndexing Profiles ({len(profiles)}):")
    for p in profiles[:20]:
        lines.append(f"  - {p['name']} → {p['target_cat'] or '(no target)'}")
    if len(profiles) > 20:
        lines.append(f"  ... and {len(profiles) - 20} more")

    # eForms
    ef_list = [e for e in (eforms or []) if e.get("name")]
    if ef_list:
        lines.append(f"\neForms ({len(ef_list)}):")
        for ef in ef_list[:15]:
            lines.append(f"  - {ef['name']}")
        if len(ef_list) > 15:
            lines.append(f"  ... and {len(ef_list) - 15} more")

    # Keyword dictionaries
    kw = maps.get("kw_dicts", {})
    if kw:
        lines.append(f"\nKeyword Dictionaries ({len(kw)}): {', '.join(list(kw.keys())[:10])}")

    # Retention policies
    # (passed via server_info or maps — skip if not present)

    lines += ["", "Write the executive summary now:"]
    return "\n".join(lines)


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

    kwargs = {"model": ai_model or "local-model", "messages": [{"role": "user", "content": prompt}],
              "temperature": 0.4, "max_tokens": 1024}

    response = client.chat.completions.create(**kwargs)
    text = response.choices[0].message.content.strip()

    if log_fn:
        log_fn("AI summary received.")

    return text
