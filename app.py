import sys
import os
import json
import time
import argparse
import threading
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd

from powerbi_chatbot import (
    acquire_token,
    execute_dax,
    result_to_dataframe,
    generate_response,
    build_dax_instruction,
    build_answer_instruction,
    parse_json_strict,
    _extract_pbi_error,
    begin_question,
    flush_tokens,
    DAX_MAX_RETRIES,
    SCHEMA_PATH,
)

# ── Power BI report embed ─────────────────────────────────────────────────────
# Reporte "Natco Executive Sales Dashboard". Se puede sobreescribir con la variable
# de entorno POWERBI_EMBED_URL. Si queda vacía, se muestra un placeholder.
POWERBI_EMBED_URL = os.environ.get(
    "POWERBI_EMBED_URL",
    "https://app.powerbi.com/reportEmbed"
    "?reportId=03f15a97-af85-4a62-b681-bcf79c975c3f"
    "&autoAuth=true"
    "&ctid=fe373213-5789-445c-ad21-62e1d031e688",
)
POWERBI_REPORT_TITLE = os.environ.get("POWERBI_REPORT_TITLE", "Natco Executive Sales Dashboard")

# ── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PBI Analyst",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Custom CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@300;400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');

/* ── Design tokens (light) ──
   bg      #ffffff   app background
   surface #f5f7fa   raised surfaces / assistant bubble
   line    #e3e8ef   borders
   text    #1b2330   primary text (near-black slate)
   muted   #5a6573   secondary text
   faint   #8a93a1   tertiary / labels
   gold    #f5c518   primary accent (Power BI) — used as button bg w/ dark text
   teal    #0b7d74   secondary accent (data / code), readable on white
*/

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    background-color: #ffffff;
    color: #1b2330;
}
.stApp { background-color: #ffffff; }

#MainMenu, footer, header { visibility: hidden; }
.stDeployButton { display: none; }
[data-testid="stToolbar"] { display: none; }

.block-container { padding-top: 1.4rem; max-width: 1500px; }

/* ── Header ── */
.pbi-header {
    display: flex;
    align-items: center;
    gap: 14px;
    padding: 6px 0 18px 0;
    border-bottom: 1px solid #e3e8ef;
    margin-bottom: 22px;
}
.pbi-logo {
    width: 40px; height: 40px;
    background: linear-gradient(135deg, #f5c518 0%, #e0a800 100%);
    border-radius: 10px;
    display: flex; align-items: center; justify-content: center;
    font-size: 21px; flex-shrink: 0;
    box-shadow: 0 4px 14px #f5c51833;
}
.pbi-title {
    font-family: 'Space Grotesk', sans-serif;
    font-size: 23px; font-weight: 600;
    color: #111722; letter-spacing: -0.4px;
}
.pbi-subtitle {
    font-size: 11px; color: #8a93a1;
    font-family: 'JetBrains Mono', monospace;
    letter-spacing: 1px; text-transform: uppercase; margin-top: 3px;
}
.status-dot {
    width: 9px; height: 9px; border-radius: 50%;
    margin-left: 6px; flex-shrink: 0;
}
.status-dot.connected { background: #2da44e; box-shadow: 0 0 8px #2da44e55; }
.status-dot.disconnected { background: #d1242f; }
.status-dot.connecting { background: #e0a800; animation: pulse 1.5s ease-in-out infinite; }
@keyframes pulse { 0%,100%{opacity:1;} 50%{opacity:.3;} }

/* ── Pane titles ── */
.pane-label {
    font-family: 'JetBrains Mono', monospace;
    font-size: 11px; letter-spacing: 1.5px; text-transform: uppercase;
    color: #5a6573; margin: 0 0 12px 2px;
    display: flex; align-items: center; gap: 8px;
}
.pane-label::before {
    content: ""; width: 14px; height: 2px;
    background: #f5c518; border-radius: 2px;
}

/* ── Report frame ── */
.report-frame {
    border: 1px solid #e3e8ef; border-radius: 14px;
    overflow: hidden; background: #ffffff;
}
.report-placeholder {
    border: 1px dashed #d4dbe4; border-radius: 14px;
    background: #f5f7fa; height: 700px;
    display: flex; flex-direction: column; align-items: center; justify-content: center;
    text-align: center; padding: 40px;
}

/* ── Chat messages ── */
.msg-wrapper {
    display: flex; flex-direction: column; gap: 5px;
    margin-bottom: 20px; animation: fadeIn 0.3s ease;
}
@keyframes fadeIn { from{opacity:0;transform:translateY(6px);} to{opacity:1;transform:translateY(0);} }
.msg-role {
    font-size: 10px; font-family: 'JetBrains Mono', monospace;
    letter-spacing: 1px; text-transform: uppercase; color: #7d8896; padding-left: 2px;
}
.msg-bubble {
    padding: 13px 17px; border-radius: 13px;
    font-size: 14.5px; line-height: 1.65; max-width: 88%;
}
.msg-bubble.user {
    background: #f5c518; color: #1a1407; font-weight: 500;
    align-self: flex-end; border-bottom-right-radius: 4px;
}
.msg-bubble.assistant {
    background: #f5f7fa; border: 1px solid #e3e8ef; color: #1b2330;
    align-self: flex-start; border-bottom-left-radius: 4px;
}
.msg-bubble.error {
    background: #fdeceb; border: 1px solid #f5c2bd; color: #b42318;
    align-self: flex-start;
}

/* ── DAX block ── */
.dax-block {
    background: #f4f6f9; border: 1px solid #e3e8ef; border-radius: 10px;
    padding: 14px 16px; font-family: 'JetBrains Mono', monospace;
    font-size: 12.5px; color: #0b6b63; line-height: 1.7;
    overflow-x: auto; white-space: pre; margin-top: 10px;
}
.dax-label, .data-label {
    font-size: 10px; font-family: 'JetBrains Mono', monospace;
    letter-spacing: 1px; text-transform: uppercase; color: #5a6573; margin-bottom: 6px;
}

/* ── Tool call badge ── */
.tool-badge {
    display: inline-flex; align-items: center; gap: 6px;
    background: #e7f4f2; border: 1px solid #bfe3df; border-radius: 20px;
    padding: 4px 12px; font-size: 11px;
    font-family: 'JetBrains Mono', monospace; color: #0b7d74; margin: 4px 4px 4px 0;
}

/* ── Notes ── */
.notes-text { font-size: 12.5px; color: #5a6573; font-style: italic; margin-top: 8px; padding-left: 2px; }

/* ── Dataframe / result table ── */
[data-testid="stDataFrame"] {
    border: 1px solid #e3e8ef !important; border-radius: 10px !important; overflow: hidden;
}
.result-wrap {
    margin-top: 8px; max-height: 340px; overflow: auto;
    border: 1px solid #e3e8ef; border-radius: 10px;
}
table.result-table {
    width: 100%; border-collapse: collapse; font-size: 13px;
    font-family: 'Inter', sans-serif; color: #1b2330; background: #ffffff;
}
table.result-table thead th {
    position: sticky; top: 0; background: #f1f4f8; color: #5a6573;
    font-family: 'JetBrains Mono', monospace; font-size: 10.5px;
    letter-spacing: 0.6px; text-transform: uppercase; font-weight: 600;
    text-align: left; padding: 10px 14px; border-bottom: 1px solid #e3e8ef;
}
table.result-table tbody td {
    padding: 9px 14px; border-bottom: 1px solid #eef1f5; color: #1b2330;
}
table.result-table tbody tr:last-child td { border-bottom: none; }
table.result-table tbody tr:nth-child(even) { background: #fafbfc; }

/* ── Text input ── */
.stTextInput > div > div > input {
    background: #ffffff !important; border: 1px solid #d4dbe4 !important;
    border-radius: 11px !important; color: #1b2330 !important;
    font-family: 'Inter', sans-serif !important; font-size: 14px !important;
    padding: 13px 16px !important; caret-color: #c79100;
}
.stTextInput > div > div > input:focus {
    border-color: #f5c518 !important; box-shadow: 0 0 0 3px #f5c5182e !important;
}
.stTextInput > div > div > input::placeholder { color: #8a93a1 !important; }

/* ── Buttons — shared shape ── */
.stButton > button, .stFormSubmitButton > button {
    border-radius: 10px !important; font-family: 'Space Grotesk', sans-serif !important;
    font-weight: 600 !important; font-size: 13.5px !important;
    padding: 11px 18px !important; letter-spacing: 0.2px;
    transition: all 0.15s ease !important; height: 48px !important;
    white-space: nowrap !important; min-width: 0 !important;
}
/* Primary (gold) — Send + Sign in */
button[kind="primary"], button[kind="primaryFormSubmit"] {
    background: #f5c518 !important; color: #1a1407 !important; border: none !important;
}
button[kind="primary"]:hover, button[kind="primaryFormSubmit"]:hover {
    background: #ffd83a !important; transform: translateY(-1px);
    box-shadow: 0 5px 16px #f5c51840 !important;
}
/* Secondary (outline) — Dashboard toggle */
button[kind="secondary"], button[kind="secondaryFormSubmit"] {
    background: #ffffff !important; color: #1b2330 !important;
    border: 1px solid #d4dbe4 !important;
}
button[kind="secondary"]:hover, button[kind="secondaryFormSubmit"]:hover {
    border-color: #f5c518 !important; color: #946d00 !important;
    background: #fffbe9 !important;
}
.stButton > button:active, .stFormSubmitButton > button:active { transform: translateY(0) !important; }

/* form has no border/padding box */
[data-testid="stForm"] { border: none !important; padding: 0 !important; }

/* ── Dashboard modal — large & light ── */
div[data-testid="stDialog"] div[role="dialog"] {
    width: 96vw !important; max-width: 1700px !important; max-height: 94vh !important;
    background: #ffffff !important; border: 1px solid #e3e8ef !important;
}
/* Dialog title — force dark & bold (default is light gray under a dark base theme) */
div[data-testid="stDialog"] h1,
div[data-testid="stDialog"] h2,
div[data-testid="stDialog"] h3,
div[data-testid="stDialog"] [data-testid="stHeading"],
div[data-testid="stDialog"] [role="dialog"] > div:first-child {
    color: #111722 !important; opacity: 1 !important;
    font-family: 'Space Grotesk', sans-serif !important; font-weight: 600 !important;
}

/* ── Spinner — dark text so "Thinking…" is readable on white ── */
[data-testid="stSpinner"] { color: #1b2330 !important; }
[data-testid="stSpinner"] p { color: #1b2330 !important; font-family: 'Inter', sans-serif !important; }
.stSpinner > div { border-top-color: #f5c518 !important; }

/* ── Expander — force light even if Streamlit base theme is dark ── */
[data-testid="stExpander"] {
    background: #f9fafc !important; border: 1px solid #e3e8ef !important; border-radius: 10px !important;
}
[data-testid="stExpander"] details,
[data-testid="stExpander"] [data-testid="stExpanderDetails"] { background: #f9fafc !important; }
[data-testid="stExpander"] summary {
    background: #f9fafc !important; color: #5a6573 !important; font-size: 12px !important;
    font-family: 'JetBrains Mono', monospace !important;
}
[data-testid="stExpander"] summary:hover { color: #1b2330 !important; }
[data-testid="stExpander"] summary svg { fill: #5a6573 !important; }

/* ── Auth screen ── */
.auth-card {
    max-width: 460px; margin: 70px auto; background: #ffffff;
    border: 1px solid #e3e8ef; border-radius: 18px; padding: 44px; text-align: center;
    box-shadow: 0 8px 30px #1b23300d;
}
.auth-icon { font-size: 48px; margin-bottom: 20px; }
.auth-title {
    font-family: 'Space Grotesk', sans-serif; font-size: 23px; font-weight: 600;
    color: #111722; margin-bottom: 10px;
}
.auth-desc { font-size: 13.5px; color: #5a6573; line-height: 1.6; margin-bottom: 4px; }

/* ── Token counter ── */
.token-counter { display: flex; gap: 18px; margin-left: auto; align-items: center; }
.token-stat { text-align: right; }
.token-stat-value {
    font-size: 14px; font-family: 'JetBrains Mono', monospace; color: #946d00; font-weight: 600;
}
.token-stat-label {
    font-size: 9px; font-family: 'JetBrains Mono', monospace;
    letter-spacing: 0.8px; text-transform: uppercase; color: #8a93a1;
}
.token-divider { width: 1px; height: 30px; background: #e3e8ef; }

::-webkit-scrollbar { width: 8px; height: 8px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: #cfd6df; border-radius: 4px; }
::-webkit-scrollbar-thumb:hover { background: #b6bfca; }
</style>
""", unsafe_allow_html=True)


# ── Auth method from CLI args ────────────────────────────────────────────────
def _get_auth_method() -> str:
    """Parse --auth from sys.argv (Streamlit passes unknown args through)."""
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--auth", default="auto",
                        choices=["auto", "env", "powershell", "devicecode"])
    args, _ = parser.parse_known_args()
    return args.auth


# ── Session state init ───────────────────────────────────────────────────────
if "token" not in st.session_state:
    st.session_state.token = None
if "schema" not in st.session_state:
    st.session_state.schema = None
if "messages" not in st.session_state:
    st.session_state.messages = []
if "auth_error" not in st.session_state:
    st.session_state.auth_error = None
if "pending_question" not in st.session_state:
    st.session_state.pending_question = None
if "session_tokens" not in st.session_state:
    st.session_state.session_tokens = {"input": 0, "output": 0}
if "device_flow" not in st.session_state:
    st.session_state.device_flow = None
if "device_flow_app" not in st.session_state:
    st.session_state.device_flow_app = None
if "device_flow_result" not in st.session_state:
    st.session_state.device_flow_result = None


# ── Ask with metadata ────────────────────────────────────────────────────────
def ask_with_meta(question: str, schema: dict, token: str) -> dict:
    """
    Runs the full pipeline and returns a dict with:
      answer, dax, notes, df, tool_calls, tokens, error
    so the UI can render each piece separately.
    """
    result = {
        "answer":     "",
        "dax":        "",
        "notes":      "",
        "df":         None,
        "tool_calls": [],
        "tokens":     {"input": 0, "output": 0},
        "error":      None,
    }

    begin_question(question)
    prior_dax = ""
    error_msg = ""
    success   = False

    try:
        for attempt in range(1, DAX_MAX_RETRIES + 1):
            if attempt == 1:
                raw, tool_calls = _run_fc_turn_with_log(question, schema, token)
                result["tool_calls"] = tool_calls
            else:
                raw = generate_response(
                    build_dax_instruction(schema, question, prior_dax, error_msg)
                )

            try:
                obj = parse_json_strict(raw)
            except (ValueError, Exception) as exc:
                result["error"] = f"Could not parse model output: {exc}\n\n{raw[:400]}"
                return result

            dax   = obj.get("dax", "").strip()
            notes = obj.get("notes", "")
            result["dax"]   = dax
            result["notes"] = notes

            if "EVALUATE" not in dax.upper():
                result["error"] = f"Generated DAX contains no EVALUATE statement:\n{dax[:300]}"
                return result

            try:
                result_json = execute_dax(token, dax)
                break
            except RuntimeError as exc:
                error_msg = _extract_pbi_error(exc)
                prior_dax = dax
                if attempt == DAX_MAX_RETRIES:
                    result["error"] = (f"DAX failed after {DAX_MAX_RETRIES} attempts.\n\n"
                                       f"Last error: {error_msg}")
                    return result
                continue

        df = result_to_dataframe(result_json)
        if df.empty:
            result["answer"] = "The query returned no rows."
            success = True
            return result

        result["df"]     = df
        result["answer"] = generate_response(build_answer_instruction(question, df))
        success = True
        return result

    finally:
        counts = flush_tokens(success=success)
        result["tokens"] = counts


def _run_fc_turn_with_log(question, schema, token):
    """
    Runs the function-calling turn and captures which tool calls were made.
    Returns (raw_text, list_of_tool_call_strings).
    """
    import powerbi_chatbot as _bot

    tool_calls_log = []
    messages = [{"role": "user", "content": _bot.build_dax_instruction(schema, question)}]

    while True:
        response   = _bot._call_openai(messages, use_tools=True)
        msg        = response.choices[0].message
        tool_calls = msg.tool_calls or []

        if not tool_calls:
            return msg.content or "", tool_calls_log

        # Append GPT's response to conversation
        messages.append(msg)

        for tc in tool_calls:
            fn_name = tc.function.name
            try:
                args = json.loads(tc.function.arguments)
            except json.JSONDecodeError:
                args = {}

            table  = args.get("table", "")
            column = args.get("column", "")
            hint   = args.get("search_hint", question)

            tool_calls_log.append(f"{fn_name}({table}[{column}], hint='{hint}')")

            if fn_name == "get_column_values" and table and column:
                values  = _bot.fetch_column_values(token, table, column, hint)
                payload = {"values": values}
            else:
                payload = {"error": f"Unknown function or missing args: {fn_name}"}

            messages.append({
                "role":         "tool",
                "tool_call_id": tc.id,
                "content":      json.dumps(payload),
            })


# ── Render a single chat message ─────────────────────────────────────────────
def render_message(msg: dict):
    role    = msg["role"]
    content = msg["content"]
    meta    = msg.get("meta", {})

    role_label = "You" if role == "user" else "PBI Analyst"
    bubble_cls = "user" if role == "user" else "assistant"

    if meta.get("error"):
        bubble_cls = "error"

    st.markdown(f"""
    <div class="msg-wrapper">
        <div class="msg-role">{role_label}</div>
        <div class="msg-bubble {bubble_cls}">{content}</div>
    </div>
    """, unsafe_allow_html=True)

    # Tool calls badge row
    if meta.get("tool_calls"):
        badges = "".join(
            f'<span class="tool-badge">⚡ {tc}</span>'
            for tc in meta["tool_calls"]
        )
        st.markdown(badges, unsafe_allow_html=True)

    # DAX + data in expander
    if meta.get("dax") and not meta.get("error"):
        with st.expander("View DAX & data", expanded=False):
            if meta.get("notes"):
                st.markdown(f'<div class="notes-text">📝 {meta["notes"]}</div>',
                            unsafe_allow_html=True)

            # Token usage for this message
            if meta.get("tokens"):
                t = meta["tokens"]
                total = t["input"] + t["output"]
                st.markdown(
                    f'<div class="notes-text" style="color:#5a6573;">🔢 Tokens — '
                    f'in: <span style="color:#946d00;font-weight:600">{t["input"]:,}</span> · '
                    f'out: <span style="color:#946d00;font-weight:600">{t["output"]:,}</span> · '
                    f'total: <span style="color:#946d00;font-weight:600">{total:,}</span></div>',
                    unsafe_allow_html=True,
                )
            st.markdown('<div class="dax-label">Generated DAX</div>',
                        unsafe_allow_html=True)
            st.markdown(f'<div class="dax-block">{meta["dax"]}</div>',
                        unsafe_allow_html=True)

            if meta.get("df") is not None and not meta["df"].empty:
                st.markdown('<div class="data-label" style="margin-top:16px;">Result table</div>',
                            unsafe_allow_html=True)
                df    = meta["df"]
                shown = df.head(200)
                table_html = shown.to_html(index=False, border=0, classes="result-table")
                more = ("" if len(df) <= 200
                        else f'<div class="notes-text">… {len(df) - 200:,} more rows</div>')
                st.markdown(f'<div class="result-wrap">{table_html}</div>{more}',
                            unsafe_allow_html=True)


# ── Auth helpers ──────────────────────────────────────────────────────────────
def _command_exists(cmd: str) -> bool:
    import shutil
    return shutil.which(cmd) is not None


def _start_device_code_flow():
    """
    Starts MSAL device-code flow using the well-known Power BI public client.
    Stores the flow and app object in session state for polling in the main thread.
    No app registration required — same client the PowerShell module uses.
    """
    import msal as _msal

    POWERBI_PUBLIC_CLIENT_ID = "ea0616ba-638b-4df5-95b9-636659ae5121"
    app = _msal.PublicClientApplication(
        client_id=POWERBI_PUBLIC_CLIENT_ID,
        authority="https://login.microsoftonline.com/organizations",
    )
    flow = app.initiate_device_flow(
        scopes=["https://analysis.windows.net/powerbi/api/.default"]
    )
    if "user_code" not in flow:
        raise RuntimeError(f"Could not start device-code flow: {flow}")

    # Store both app and flow — we'll poll from the main Streamlit thread
    st.session_state.device_flow        = flow
    st.session_state.device_flow_app    = app
    st.session_state.device_flow_result = "pending"


# ── Auth screen ───────────────────────────────────────────────────────────────
def render_auth_screen():
    st.markdown("""
    <div class="auth-card">
        <div class="auth-icon">🔐</div>
        <div class="auth-title">Connect to Microsoft Fabric</div>
        <div class="auth-desc">
            Sign in with your Microsoft account to start querying your
            Fabric dataset in natural language.
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Env token in Streamlit secrets — silent auto-login ────────────────────
    env_token = os.environ.get("POWERBI_TOKEN")
    if env_token:
        try:
            schema = json.load(open(SCHEMA_PATH, encoding="utf-8"))
            st.session_state.token  = env_token
            st.session_state.schema = schema
            st.rerun()
        except Exception as exc:
            st.error(f"Token found in secrets but failed to load schema: {exc}")
        return

    # ── Device-code token arrived ─────────────────────────────────────────────
    flow_result = st.session_state.get("device_flow_result")
    flow        = st.session_state.get("device_flow")

    if flow_result and flow_result not in ("pending",) \
            and not flow_result.startswith("error:"):
        try:
            schema = json.load(open(SCHEMA_PATH, encoding="utf-8"))
            st.session_state.token              = flow_result
            st.session_state.schema             = schema
            st.session_state.device_flow        = None
            st.session_state.device_flow_result = None
            st.rerun()
        except Exception as exc:
            st.session_state.auth_error = str(exc)
        return

    # ── Device-code error ─────────────────────────────────────────────────────
    if flow_result and flow_result.startswith("error:"):
        st.error("Sign-in failed: " + flow_result.removeprefix("error:"))
        st.session_state.device_flow        = None
        st.session_state.device_flow_result = None

    # ── Waiting for user to complete sign-in ──────────────────────────────────
    if flow and flow_result == "pending":
        st.markdown(f"""
        <div style="max-width:460px;margin:0 auto;background:#ffffff;
                    border:1px solid #e3e8ef;border-radius:14px;
                    padding:32px;text-align:center;box-shadow:0 8px 30px #1b23300d;">
            <div style="font-size:13px;color:#5a6573;margin-bottom:20px;line-height:1.6;">
                Open
                <a href="https://microsoft.com/devicelogin" target="_blank"
                   style="color:#946d00;text-decoration:none;font-weight:600;">
                   microsoft.com/devicelogin</a>
                and enter this code:
            </div>
            <div style="font-family:'JetBrains Mono',monospace;font-size:36px;
                        font-weight:600;color:#111722;letter-spacing:8px;
                        background:#f4f6f9;padding:18px 28px;border-radius:10px;
                        border:1px solid #e3e8ef;display:inline-block;
                        margin-bottom:24px;">
                {flow["user_code"]}
            </div>
            <div style="font-size:12px;color:#8a93a1;">
                Waiting for sign-in… refreshing automatically.
            </div>
        </div>
        """, unsafe_allow_html=True)

        app = st.session_state.get("device_flow_app")
        if app:
            try:
                result = app.acquire_token_by_device_flow(
                    {**flow, "expires_in": 4}  # short poll window
                )
                if "access_token" in result:
                    st.session_state.device_flow_result = result["access_token"]
                elif result.get("error") == "authorization_pending":
                    pass  # still waiting — rerun and show code again
                else:
                    st.session_state.device_flow_result = (
                        f"error:{result.get('error_description', result.get('error', 'Unknown'))}"
                    )
            except Exception:
                pass  # timeout or pending — just rerun

        time.sleep(2)
        st.rerun()
        return

    # ── Initial button ────────────────────────────────────────────────────────
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        if st.button("Sign in with Microsoft", use_container_width=True, type="primary"):
            if _command_exists("powershell") or _command_exists("pwsh"):
                with st.spinner("Opening browser for sign-in…"):
                    try:
                        token  = acquire_token("powershell")
                        schema = json.load(open(SCHEMA_PATH, encoding="utf-8"))
                        st.session_state.token  = token
                        st.session_state.schema = schema
                        st.rerun()
                    except Exception as exc:
                        st.session_state.auth_error = str(exc)
            else:
                try:
                    _start_device_code_flow()
                    st.rerun()
                except Exception as exc:
                    st.session_state.auth_error = str(exc)

    if st.session_state.auth_error:
        st.error(f"Authentication failed: {st.session_state.auth_error}")


# ── Power BI dashboard (slide-in modal) ───────────────────────────────────────
@st.dialog(POWERBI_REPORT_TITLE, width="large")
def show_dashboard_dialog():
    if POWERBI_EMBED_URL:
        components.html(
            f"""
            <div style="width:100%; height:800px;">
              <iframe
                  title="{POWERBI_REPORT_TITLE}"
                  src="{POWERBI_EMBED_URL}"
                  style="border:none; display:block; width:100%; height:800px;"
                  frameborder="0"
                  allowFullScreen="true">
              </iframe>
            </div>
            """,
            height=800,
            scrolling=False,
        )
    else:
        st.markdown("""
        <div class="report-placeholder" style="height:420px;">
            <div style="font-size:40px;margin-bottom:14px;">📊</div>
            <div style="font-size:15px;color:#1b2330;margin-bottom:8px;">
                No report URL configured
            </div>
            <div style="font-size:12.5px;color:#5a6573;line-height:1.7;max-width:340px;">
                Set <code style="color:#946d00;">POWERBI_EMBED_URL</code> (env var) or edit the
                constant at the top of <code style="color:#946d00;">app.py</code> with your
                report's embed link.
            </div>
        </div>
        """, unsafe_allow_html=True)


# ── Conversation ──────────────────────────────────────────────────────────────
def render_chat_pane(schema, token):
    # Dashboard toggle — right-aligned, opens the report in a modal
    _, btn_col = st.columns([3, 1])
    with btn_col:
        if st.button("📊  Open Dashboard", use_container_width=True, type="secondary",
                     help="Abrir el dashboard de Power BI"):
            show_dashboard_dialog()

    # Scrollable history
    history_box = st.container(height=520)
    with history_box:
        if not st.session_state.messages:
            st.markdown("""
            <div style="text-align:center; padding: 90px 0; color:#8a93a1;">
                <div style="font-size:34px;margin-bottom:14px;">💬</div>
                <div style="font-size:14.5px;color:#5a6573;">
                    Ask anything about your Power BI data
                </div>
                <div style="font-size:12px;color:#8a93a1;margin-top:8px;
                            font-family:'JetBrains Mono',monospace;">
                    "Top 10 customers by net sales" · "Compare Walmart Q1 2024 vs 2026"
                </div>
            </div>
            """, unsafe_allow_html=True)
        for msg in st.session_state.messages:
            render_message(msg)

    # Input row — wrapped in a form so Enter submits
    with st.form(key="ask_form", clear_on_submit=True):
        col_input, col_btn = st.columns([5, 1.4])
        with col_input:
            question = st.text_input(
                label="question",
                label_visibility="collapsed",
                placeholder="Ask a question…  (press Enter to send)",
                key="question_input",
            )
        with col_btn:
            send = st.form_submit_button("Send", use_container_width=True, type="primary")

    if send and question.strip():
        st.session_state.pending_question = question.strip()
        st.rerun()

    # Process the pending question on the rerun after submit
    if st.session_state.pending_question:
        q = st.session_state.pending_question
        st.session_state.pending_question = None

        st.session_state.messages.append({"role": "user", "content": q})

        with st.spinner("Thinking…"):
            result = ask_with_meta(q, schema, token)

        if result["error"]:
            content = f"❌ {result['error']}"
            meta    = {"error": True}
        else:
            content = result["answer"]
            meta    = {
                "dax":        result["dax"],
                "notes":      result["notes"],
                "df":         result["df"],
                "tool_calls": result["tool_calls"],
                "tokens":     result["tokens"],
            }
            st.session_state.session_tokens["input"]  += result["tokens"]["input"]
            st.session_state.session_tokens["output"] += result["tokens"]["output"]

        st.session_state.messages.append({
            "role": "assistant",
            "content": content,
            "meta": meta,
        })
        st.rerun()


# ── Main chat UI ──────────────────────────────────────────────────────────────
def render_chat():
    schema = st.session_state.schema
    token  = st.session_state.token

    # Header (full width)
    sess_in  = st.session_state.session_tokens["input"]
    sess_out = st.session_state.session_tokens["output"]
    st.markdown(f"""
    <div class="pbi-header">
        <div class="pbi-logo">📊</div>
        <div>
            <div class="pbi-title">PBI Analyst</div>
            <div class="pbi-subtitle">NATCO</div>
        </div>
        <div class="token-counter">
            <div class="token-stat">
                <div class="token-stat-value">{sess_in:,}</div>
                <div class="token-stat-label">in tokens</div>
            </div>
            <div class="token-divider"></div>
            <div class="token-stat">
                <div class="token-stat-value">{sess_out:,}</div>
                <div class="token-stat-label">out tokens</div>
            </div>
            <div class="token-divider"></div>
            <div class="token-stat">
                <div class="token-stat-value">{sess_in + sess_out:,}</div>
                <div class="token-stat-label">total</div>
            </div>
        </div>
        <div class="status-dot connected" title="Connected"></div>
    </div>
    """, unsafe_allow_html=True)

    # Clean, centered chatbot. Dashboard opens on demand from the "Tablero" button.
    _, mid, _ = st.columns([1, 3, 1])
    with mid:
        render_chat_pane(schema, token)


# ── Entry point ───────────────────────────────────────────────────────────────
def main():
    if st.session_state.token is None:
        render_auth_screen()
    else:
        render_chat()


main()
