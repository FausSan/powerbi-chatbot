import sys
import os
import json
import time
import argparse
import threading
import streamlit as st
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
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap');

/* ── Base ── */
html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
    background-color: #0a0e14;
    color: #c9d1d9;
}

.stApp {
    background-color: #0a0e14;
}

/* ── Hide streamlit chrome ── */
#MainMenu, footer, header { visibility: hidden; }
.stDeployButton { display: none; }
[data-testid="stToolbar"] { display: none; }

/* ── Header ── */
.pbi-header {
    display: flex;
    align-items: center;
    gap: 14px;
    padding: 28px 0 20px 0;
    border-bottom: 1px solid #1e2940;
    margin-bottom: 32px;
}
.pbi-logo {
    width: 38px;
    height: 38px;
    background: linear-gradient(135deg, #f2c811 0%, #e8a000 100%);
    border-radius: 8px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 20px;
    flex-shrink: 0;
}
.pbi-title {
    font-size: 22px;
    font-weight: 600;
    color: #e6edf3;
    letter-spacing: -0.3px;
}
.pbi-subtitle {
    font-size: 12px;
    color: #6e7681;
    font-family: 'IBM Plex Mono', monospace;
    letter-spacing: 0.5px;
    text-transform: uppercase;
    margin-top: 2px;
}
.status-dot {
    width: 8px;
    height: 8px;
    border-radius: 50%;
    margin-left: auto;
    flex-shrink: 0;
}
.status-dot.connected { background: #3fb950; box-shadow: 0 0 8px #3fb95066; }
.status-dot.disconnected { background: #f85149; }
.status-dot.connecting {
    background: #f2c811;
    animation: pulse 1.5s ease-in-out infinite;
}
@keyframes pulse {
    0%, 100% { opacity: 1; }
    50% { opacity: 0.3; }
}

/* ── Chat messages ── */
.msg-wrapper {
    display: flex;
    flex-direction: column;
    gap: 4px;
    margin-bottom: 24px;
    animation: fadeIn 0.3s ease;
}
@keyframes fadeIn {
    from { opacity: 0; transform: translateY(6px); }
    to   { opacity: 1; transform: translateY(0); }
}
.msg-role {
    font-size: 10px;
    font-family: 'IBM Plex Mono', monospace;
    letter-spacing: 1px;
    text-transform: uppercase;
    color: #6e7681;
    padding-left: 2px;
}
.msg-bubble {
    padding: 14px 18px;
    border-radius: 10px;
    font-size: 14.5px;
    line-height: 1.65;
    max-width: 85%;
}
.msg-bubble.user {
    background: #161b27;
    border: 1px solid #1e2940;
    color: #c9d1d9;
    align-self: flex-end;
    border-bottom-right-radius: 3px;
}
.msg-bubble.assistant {
    background: #0d1117;
    border: 1px solid #1e2940;
    color: #c9d1d9;
    align-self: flex-start;
    border-bottom-left-radius: 3px;
}
.msg-bubble.error {
    background: #1a0a0a;
    border: 1px solid #f8514933;
    color: #f85149;
}

/* ── DAX expander ── */
.dax-block {
    background: #060a10;
    border: 1px solid #1e2940;
    border-radius: 8px;
    padding: 14px 16px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 12.5px;
    color: #79c0ff;
    line-height: 1.7;
    overflow-x: auto;
    white-space: pre;
    margin-top: 10px;
}
.dax-label {
    font-size: 10px;
    font-family: 'IBM Plex Mono', monospace;
    letter-spacing: 1px;
    text-transform: uppercase;
    color: #6e7681;
    margin-bottom: 6px;
}

/* ── Tool call badge ── */
.tool-badge {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    background: #0d1b2a;
    border: 1px solid #1e3a5f;
    border-radius: 20px;
    padding: 4px 12px;
    font-size: 11px;
    font-family: 'IBM Plex Mono', monospace;
    color: #58a6ff;
    margin: 4px 4px 4px 0;
}

/* ── Data table ── */
.data-section {
    margin-top: 14px;
}
.data-label {
    font-size: 10px;
    font-family: 'IBM Plex Mono', monospace;
    letter-spacing: 1px;
    text-transform: uppercase;
    color: #6e7681;
    margin-bottom: 8px;
}

/* ── Streamlit dataframe overrides ── */
[data-testid="stDataFrame"] {
    border: 1px solid #1e2940 !important;
    border-radius: 8px !important;
    overflow: hidden;
}

/* ── Input area ── */
.input-area {
    position: sticky;
    bottom: 0;
    background: linear-gradient(transparent, #0a0e14 30%);
    padding: 20px 0 16px 0;
    margin-top: 12px;
}

/* ── Streamlit input overrides ── */
.stTextInput > div > div > input {
    background: #0d1117 !important;
    border: 1px solid #30363d !important;
    border-radius: 10px !important;
    color: #e6edf3 !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
    font-size: 14px !important;
    padding: 12px 16px !important;
    caret-color: #f2c811;
}
.stTextInput > div > div > input:focus {
    border-color: #f2c811 !important;
    box-shadow: 0 0 0 3px #f2c81118 !important;
}
.stTextInput > div > div > input::placeholder {
    color: #484f58 !important;
}

/* ── Button ── */
.stButton > button {
    background: #f2c811 !important;
    color: #0a0e14 !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
    font-weight: 600 !important;
    font-size: 13px !important;
    padding: 10px 22px !important;
    letter-spacing: 0.3px;
    transition: all 0.15s ease !important;
    height: 46px !important;
}
.stButton > button:hover {
    background: #ffd433 !important;
    transform: translateY(-1px);
    box-shadow: 0 4px 12px #f2c81130 !important;
}
.stButton > button:active {
    transform: translateY(0) !important;
}

/* ── Auth screen ── */
.auth-card {
    max-width: 440px;
    margin: 80px auto;
    background: #0d1117;
    border: 1px solid #1e2940;
    border-radius: 16px;
    padding: 40px;
    text-align: center;
}
.auth-icon {
    font-size: 48px;
    margin-bottom: 20px;
}
.auth-title {
    font-size: 22px;
    font-weight: 600;
    color: #e6edf3;
    margin-bottom: 8px;
}
.auth-desc {
    font-size: 13.5px;
    color: #6e7681;
    line-height: 1.6;
    margin-bottom: 28px;
}

/* ── Spinner override ── */
.stSpinner > div {
    border-top-color: #f2c811 !important;
}

/* ── Expander ── */
[data-testid="stExpander"] {
    background: #0d1117 !important;
    border: 1px solid #1e2940 !important;
    border-radius: 8px !important;
}
[data-testid="stExpander"] summary {
    color: #6e7681 !important;
    font-size: 12px !important;
    font-family: 'IBM Plex Mono', monospace !important;
}

/* ── Token counter ── */
.token-counter {
    display: flex;
    gap: 16px;
    margin-left: auto;
    align-items: center;
}
.token-stat {
    text-align: right;
}
.token-stat-value {
    font-size: 13px;
    font-family: 'IBM Plex Mono', monospace;
    color: #f2c811;
    font-weight: 500;
}
.token-stat-label {
    font-size: 9px;
    font-family: 'IBM Plex Mono', monospace;
    letter-spacing: 0.8px;
    text-transform: uppercase;
    color: #484f58;
}
.token-divider {
    width: 1px;
    height: 28px;
    background: #1e2940;
}
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: #0a0e14; }
::-webkit-scrollbar-thumb { background: #21262d; border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: #30363d; }

/* ── Notes text ── */
.notes-text {
    font-size: 12.5px;
    color: #8b949e;
    font-style: italic;
    margin-top: 8px;
    padding-left: 2px;
}
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
                    f'<div class="notes-text" style="color:#484f58;">🔢 Tokens — '
                    f'in: <span style="color:#f2c811">{t["input"]:,}</span> · '
                    f'out: <span style="color:#f2c811">{t["output"]:,}</span> · '
                    f'total: <span style="color:#f2c811">{total:,}</span></div>',
                    unsafe_allow_html=True,
                )
            st.markdown('<div class="dax-label">Generated DAX</div>',
                        unsafe_allow_html=True)
            st.markdown(f'<div class="dax-block">{meta["dax"]}</div>',
                        unsafe_allow_html=True)

            if meta.get("df") is not None and not meta["df"].empty:
                st.markdown('<div class="data-label" style="margin-top:16px;">Result table</div>',
                            unsafe_allow_html=True)
                st.dataframe(
                    meta["df"],
                    use_container_width=True,
                    hide_index=True,
                )


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
        <div style="max-width:460px;margin:0 auto;background:#0d1117;
                    border:1px solid #1e2940;border-radius:12px;
                    padding:32px;text-align:center;">
            <div style="font-size:13px;color:#8b949e;margin-bottom:20px;line-height:1.6;">
                Open
                <a href="https://microsoft.com/devicelogin" target="_blank"
                   style="color:#f2c811;text-decoration:none;font-weight:600;">
                   microsoft.com/devicelogin</a>
                and enter this code:
            </div>
            <div style="font-family:'IBM Plex Mono',monospace;font-size:36px;
                        font-weight:600;color:#e6edf3;letter-spacing:8px;
                        background:#060a10;padding:18px 28px;border-radius:8px;
                        border:1px solid #30363d;display:inline-block;
                        margin-bottom:24px;">
                {flow["user_code"]}
            </div>
            <div style="font-size:12px;color:#484f58;">
                Waiting for sign-in… refreshing automatically.
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Poll directly in the main thread with a short timeout
        # acquire_token_by_device_flow blocks until done or expired,
        # so we use the non-blocking check via initiate_device_flow's expires_in
        app = st.session_state.get("device_flow_app")
        if app:
            try:
                # Try to acquire with a 4-second window — if not ready yet,
                # MSAL raises or returns an error dict we can check
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
    # Local Windows: PowerShell browser popup (fast, seamless)
    # Streamlit Cloud / Linux: device-code shown on screen
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        if st.button("Sign in with Microsoft", use_container_width=True):
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


# ── Main chat UI ──────────────────────────────────────────────────────────────
def render_chat():
    schema = st.session_state.schema
    token  = st.session_state.token

    # Header
    sess_in  = st.session_state.session_tokens["input"]
    sess_out = st.session_state.session_tokens["output"]
    st.markdown(f"""
    <div class="pbi-header">
        <div class="pbi-logo">📊</div>
        <div>
            <div class="pbi-title">PBI Analyst</div>
            <div class="pbi-subtitle">GPT · Power BI · Natural Language</div>
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

    # Chat history
    for msg in st.session_state.messages:
        render_message(msg)

    # Empty state
    if not st.session_state.messages:
        st.markdown("""
        <div style="text-align:center; padding: 60px 0; color: #484f58;">
            <div style="font-size: 36px; margin-bottom: 16px;">💬</div>
            <div style="font-size: 15px; color: #6e7681;">
                Ask anything about your Power BI data
            </div>
            <div style="font-size: 12px; color: #484f58; margin-top: 8px; font-family: 'IBM Plex Mono', monospace;">
                e.g. "Top 10 customers by net sales" · "Compare Walmart Q1 2024 vs 2026"
            </div>
        </div>
        """, unsafe_allow_html=True)

    # Input row
    st.markdown('<div class="input-area">', unsafe_allow_html=True)
    col_input, col_btn = st.columns([6, 1])

    with col_input:
        question = st.text_input(
            label="question",
            label_visibility="collapsed",
            placeholder="Ask a question about your data…",
            key="question_input",
        )
    with col_btn:
        send = st.button("Send", use_container_width=True)

    st.markdown('</div>', unsafe_allow_html=True)

    # On Send: store the question and clear the input, then rerun once to show
    # the user bubble and trigger processing on the next pass.
    if send and question.strip():
        st.session_state.pending_question = question.strip()
        st.rerun()

    # Process the pending question (runs on the rerun after Send was clicked)
    if st.session_state.pending_question:
        q = st.session_state.pending_question
        st.session_state.pending_question = None  # clear immediately to prevent re-runs

        # Add user message to history
        st.session_state.messages.append({"role": "user", "content": q})

        # Run pipeline
        with st.spinner("Thinking…"):
            result = ask_with_meta(q, schema, token)

        # Build assistant message
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
            # Accumulate into session totals
            st.session_state.session_tokens["input"]  += result["tokens"]["input"]
            st.session_state.session_tokens["output"] += result["tokens"]["output"]

        st.session_state.messages.append({
            "role": "assistant",
            "content": content,
            "meta": meta,
        })
        st.rerun()


# ── Entry point ───────────────────────────────────────────────────────────────
def main():
    if st.session_state.token is None:
        render_auth_screen()
    else:
        render_chat()


main()
