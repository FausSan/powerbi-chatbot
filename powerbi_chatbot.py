import os
import sys
import json
import time
import re
import argparse
import subprocess
import textwrap
from typing import Any, Dict, List, Optional

import requests
import msal
import pandas as pd
from dotenv import load_dotenv
from openai import AzureOpenAI


# ──────────────────────────────────────────
# Config
# ──────────────────────────────────────────
load_dotenv()

TENANT_ID   = os.environ["TENANT_ID"]
CLIENT_ID   = os.environ["CLIENT_ID"]
GROUP_ID    = os.environ["GROUP_ID"]
DATASET_ID  = os.environ["DATASET_ID"]

AZURE_OPENAI_ENDPOINT   = os.environ["AZURE_OPENAI_ENDPOINT"]
AZURE_OPENAI_API_KEY    = os.environ["AZURE_OPENAI_API_KEY"]
AZURE_OPENAI_DEPLOYMENT = os.environ.get("AZURE_OPENAI_DEPLOYMENT", "gpt-4o")
AZURE_OPENAI_API_VERSION= os.environ.get("AZURE_OPENAI_API_VERSION", "2024-02-15-preview")
SCHEMA_PATH             = os.environ.get("SCHEMA_PATH", "schema/semantic_model.json")

POWERBI_RESOURCE = "https://analysis.windows.net/powerbi/api"
EXECUTE_DAX_URL  = (
    f"https://api.powerbi.com/v1.0/myorg/groups/{GROUP_ID}"
    f"/datasets/{DATASET_ID}/executeQueries"
)

MAX_ROWS_RETURNED = 2000   # hard cap sent to Power BI
MAX_ROWS_TO_LLM   = 1000   # rows forwarded to Gemini for summarisation


# ──────────────────────────────────────────
# Authentication
# ──────────────────────────────────────────

def _token_from_env() -> Optional[str]:
    """Read a pre-obtained token from the environment."""
    return os.environ.get("POWERBI_TOKEN") or None


def _token_from_powershell() -> str:
    """
    Equivalent of:
        Connect-PowerBIServiceAccount
        $token = Get-PowerBIAccessToken
    Requires the MicrosoftPowerBIMgmt module and PowerShell ≥ 5.
    """
    ps_exe = "pwsh" if _command_exists("pwsh") else "powershell"

    ps_script = textwrap.dedent("""
        Import-Module MicrosoftPowerBIMgmt -ErrorAction Stop
        Connect-PowerBIServiceAccount | Out-Null
        $t = Get-PowerBIAccessToken
        Write-Output $t["Authorization"]
    """)

    print("[auth] Launching PowerShell — a browser window will open for sign-in.")
    result = subprocess.run(
        # Remove -NonInteractive so the browser popup can appear
        [ps_exe, "-NoProfile", "-Command", ps_script],
        capture_output=True, text=True
    )

    if result.returncode != 0:
        raise RuntimeError(
            f"PowerShell auth failed.\nSTDERR: {result.stderr.strip()}"
        )

    lines = [l.strip() for l in result.stdout.splitlines() if l.strip()]
    if not lines:
        raise RuntimeError("PowerShell returned no token output.")

    header_value = lines[-1]
    if not header_value.startswith("Bearer "):
        raise RuntimeError(f"Unexpected token format: {header_value[:60]}")

    return header_value.split(" ", 1)[1]


def _token_from_device_code() -> str:
    """
    MSAL device-code flow using the same well-known Power BI public client
    that the MicrosoftPowerBIMgmt PowerShell module uses internally.
    No app registration required — works with any Fabric/Power BI account.
    """
    # This is the well-known Power BI Desktop / PowerShell module public client ID.
    # It has Power BI permissions pre-consented by Microsoft — no app registration needed.
    POWERBI_PUBLIC_CLIENT_ID = "ea0616ba-638b-4df5-95b9-636659ae5121"

    app = msal.PublicClientApplication(
        client_id=POWERBI_PUBLIC_CLIENT_ID,
        authority="https://login.microsoftonline.com/organizations",
    )

    flow = app.initiate_device_flow(scopes=["https://analysis.windows.net/powerbi/api/.default"])
    if "user_code" not in flow:
        raise RuntimeError(f"Device-code flow failed: {flow}")

    print(flow["message"])

    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(
            f"Device-code auth failed: {result.get('error_description', result)}"
        )

    return result["access_token"]


def _command_exists(cmd: str) -> bool:
    import shutil
    return shutil.which(cmd) is not None


def acquire_token(method: str = "auto") -> str:
    """
    method: "auto" | "env" | "powershell" | "devicecode"
    """
    if method == "auto":
        token = _token_from_env()
        if token:
            print("[auth] Using token from POWERBI_TOKEN env var.")
            return token
        if _command_exists("powershell") or _command_exists("pwsh"):
            print("[auth] PowerShell found — using PowerShell auth.")
            return _token_from_powershell()
        print("[auth] Falling back to device-code flow.")
        return _token_from_device_code()

    if method == "env":
        token = _token_from_env()
        if not token:
            raise RuntimeError("POWERBI_TOKEN env var is not set.")
        return token

    if method == "powershell":
        return _token_from_powershell()

    if method == "devicecode":
        return _token_from_device_code()

    raise ValueError(f"Unknown auth method: {method}")


# ──────────────────────────────────────────
# DAX execution
# ──────────────────────────────────────────

def execute_dax(token: str, dax: str) -> Dict[str, Any]:
    """
    POST to the executeQueries endpoint.
    https://learn.microsoft.com/en-us/rest/api/power-bi/datasets/execute-queries-in-group
    """
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    payload = {
        "queries": [{"query": dax}],
        "serializerSettings": {"includeNulls": True},
        "impersonatedUserName": None,   # omit or set to a UPN to impersonate
    }

    resp = requests.post(EXECUTE_DAX_URL, headers=headers, json=payload, timeout=120)

    # Surface a helpful error if the request fails
    if not resp.ok:
        raise RuntimeError(
            f"DAX execution failed [{resp.status_code}]: {resp.text[:500]}"
        )

    return resp.json()


def result_to_dataframe(result_json: Dict[str, Any]) -> pd.DataFrame:
    """Parse the executeQueries JSON response into a DataFrame."""
    try:
        tables = result_json["results"][0]["tables"]
    except (KeyError, IndexError):
        return pd.DataFrame()

    if not tables:
        return pd.DataFrame()

    rows = tables[0].get("rows", [])

    # Power BI prefixes column names with the table name, e.g. "Sales[Amount]"
    # Strip the prefix for cleaner output.
    df = pd.DataFrame(rows)
    df.columns = [re.sub(r"^[^\[]+\[(.+)\]$", r"\1", c) for c in df.columns]
    return df


# ──────────────────────────────────────────
# Gemini (Vertex AI) + Function Calling
# ──────────────────────────────────────────

# ──────────────────────────────────────────
# Azure OpenAI + Function Calling
# ──────────────────────────────────────────

# ── Tool definition (OpenAI function calling format) ──
OPENAI_TOOLS: List[Dict] = [
    {
        "type": "function",
        "function": {
            "name": "get_column_values",
            "description": (
                "Fetches the distinct values of any dimension column from Power BI. "
                "Call this whenever the user mentions a name, label or entity that needs "
                "to be matched to an exact stored value before writing a DAX filter. "
                "Examples: customer names, sales rep names, product names, regions, etc. "
                "Returns a JSON object with a 'values' key containing an array of strings "
                "that are the closest matches to the search hint."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "table": {
                        "type": "string",
                        "description": "The Power BI table name exactly as it appears in the schema, e.g. 'W_CUST_MBBFREP_D'.",
                    },
                    "column": {
                        "type": "string",
                        "description": "The column name exactly as it appears in the schema, e.g. 'AlphaSortName'.",
                    },
                    "search_hint": {
                        "type": "string",
                        "description": (
                            "A word or phrase from the user's question to filter the results, "
                            "e.g. 'walmart' or 'john smith'. Leave empty to return all values."
                        ),
                    },
                },
                "required": ["table", "column"],
            },
        },
    }
]

# ── Azure OpenAI client singleton ──────────
_oai_client: Optional[AzureOpenAI] = None


def _get_client() -> AzureOpenAI:
    global _oai_client
    if _oai_client is None:
        _oai_client = AzureOpenAI(
            azure_endpoint=AZURE_OPENAI_ENDPOINT,
            api_key=AZURE_OPENAI_API_KEY,
            api_version=AZURE_OPENAI_API_VERSION,
        )
    return _oai_client


def _call_openai(messages: List[Dict], use_tools: bool = False) -> Any:
    """
    Low-level Azure OpenAI call with retry logic.
    Returns the raw completion response object.
    """
    client = _get_client()
    kwargs: Dict[str, Any] = {
        "model":       AZURE_OPENAI_DEPLOYMENT,
        "messages":    messages,
        "temperature": 0.1,
        "max_tokens":  4096,
    }
    if use_tools:
        kwargs["tools"]       = OPENAI_TOOLS
        kwargs["tool_choice"] = "auto"

    for attempt in range(6):
        try:
            response = client.chat.completions.create(**kwargs)
            _accumulate_tokens(response)
            return response
        except Exception as exc:
            # Don't retry on configuration errors — fail fast with a clear message
            err_str = str(exc)
            if any(code in err_str for code in ["404", "401", "400", "NotFound", "Unauthorized", "DeploymentNotFound"]):
                raise RuntimeError(
                    f"Azure OpenAI configuration error: {exc}\n\n"
                    f"Check these .env values:\n"
                    f"  AZURE_OPENAI_ENDPOINT   = {AZURE_OPENAI_ENDPOINT}\n"
                    f"  AZURE_OPENAI_DEPLOYMENT = {AZURE_OPENAI_DEPLOYMENT}\n"
                    f"  AZURE_OPENAI_API_VERSION= {AZURE_OPENAI_API_VERSION}\n"
                    f"  AZURE_OPENAI_API_KEY    = {'set' if AZURE_OPENAI_API_KEY else 'NOT SET'}"
                ) from exc
            # Retry on rate limits and server errors
            if attempt == 5:
                raise
            wait = 2 ** attempt
            print(f"[openai] Transient error ({exc.__class__.__name__}), retrying in {wait}s…")
            time.sleep(wait)


def generate_response(instruction: str) -> str:
    """Simple single-turn text generation."""
    response = _call_openai(
        [{"role": "user", "content": instruction}],
        use_tools=False,
    )
    return response.choices[0].message.content or ""


def run_function_calling_turn(question: str, schema: Dict[str, Any], token: str) -> str:
    """
    Multi-turn Azure OpenAI conversation with function calling.

    Flow:
      1. Send the question + schema to GPT with tools available.
      2. If GPT calls get_column_values → execute it, return results, continue.
      3. Once GPT has what it needs, it produces the final DAX JSON response.

    Returns the raw text of GPT's final response (the DAX JSON string).
    """
    messages: List[Dict] = [
        {"role": "user", "content": build_dax_instruction(schema, question)}
    ]

    while True:
        response  = _call_openai(messages, use_tools=True)
        msg       = response.choices[0].message
        tool_calls = msg.tool_calls or []

        # No tool calls → GPT is done, return its text
        if not tool_calls:
            return msg.content or ""

        # Append GPT's response (with tool_calls) to the conversation
        messages.append(msg)

        # Execute each tool call and append results
        for tc in tool_calls:
            fn_name = tc.function.name
            print(f"   [tool call] {fn_name}()")

            try:
                args = json.loads(tc.function.arguments)
            except json.JSONDecodeError:
                args = {}

            if fn_name == "get_column_values":
                table       = args.get("table", "")
                column      = args.get("column", "")
                search_hint = args.get("search_hint", question)
                if table and column:
                    values  = fetch_column_values(token, table, column, search_hint)
                    payload = {"values": values}
                else:
                    payload = {"error": "Both 'table' and 'column' are required."}
            else:
                payload = {"error": f"Unknown function: {fn_name}"}

            messages.append({
                "role":         "tool",
                "tool_call_id": tc.id,
                "content":      json.dumps(payload),
            })

# ── Column value cache ────────────────────
# Shared text utilities for fuzzy matching
_STOP_WORDS = {
    "a", "an", "the", "is", "are", "was", "were", "what", "which", "who",
    "how", "when", "where", "why", "do", "does", "did", "have", "has",
    "had", "will", "would", "could", "should", "may", "might", "shall",
    "be", "been", "being", "and", "or", "but", "if", "in", "on", "at",
    "to", "for", "of", "with", "by", "from", "up", "about", "into",
    "than", "then", "so", "yet", "both", "either", "not", "no", "nor",
    "as", "just", "vs", "versus", "between", "show", "me", "get", "give",
    "list", "tell", "find", "compare", "sales", "revenue", "growth",
    "total", "top", "last", "this", "year", "month", "quarter", "q1",
    "q2", "q3", "q4", "ytd", "yoy", "2020", "2021", "2022", "2023",
    "2024", "2025", "2026", "january", "february", "march", "april",
    "may", "june", "july", "august", "september", "october", "november",
    "december",
}


def _normalize_compact(text: str) -> str:
    """Lowercase and strip ALL non-alphanumeric characters — 'WAL-MART' → 'walmart'."""
    return re.sub(r"[^a-z0-9]", "", text.lower())


def _extract_search_terms(text: str) -> list[str]:
    """
    Extract meaningful tokens from any text (question or search hint).
    Strips punctuation, lowercases, removes stop words and short tokens.
    e.g. "What are Walmart's sales Q1 2024?" → ["walmart"]
    """
    tokens = re.findall(r"[a-zA-Z0-9]+", text.lower())
    return [t for t in tokens if t not in _STOP_WORDS and len(t) >= 3]


# Stores fetched distinct values per (table, column) pair — fetched at most once
# per session each.
_column_cache: Dict[tuple, list] = {}


def fetch_column_values(token: str, table: str, column: str,
                        search_hint: str = "") -> list:
    """
    Fetch distinct values for any table/column from Power BI (cached per session).
    If search_hint is provided, returns only values whose normalized form contains
    at least one meaningful word from the hint.
    """
    cache_key = (table, column)

    if cache_key not in _column_cache:
        print(f"   [tool] Fetching distinct values for {table}[{column}]…")
        if table == "W_CUST_MBBFREP_D":
            dax = (
                f"EVALUATE CALCULATETABLE("
                f"VALUES('{table}'[{column}]), "
                f"'W_CUST_MBBFREP_D'[flag_corporate_customer] = \"Y\")"
            )
        else:
            dax = f"EVALUATE DISTINCT('{table}'[{column}])"
        try:
            result = execute_dax(token, dax)
            df = result_to_dataframe(result)
            values = df.iloc[:, 0].dropna().astype(str).tolist() if not df.empty else []
        except Exception as exc:
            print(f"   [tool] Warning: could not fetch {table}[{column}]: {exc}")
            values = []
        _column_cache[cache_key] = sorted(values)
        print(f"   [tool] {len(_column_cache[cache_key])} values loaded into cache.")

    all_values = _column_cache[cache_key]

    if not search_hint:
        return all_values

    # Extract meaningful tokens from the hint using the same stop-word filter
    terms = _extract_search_terms(search_hint)
    if not terms:
        return all_values

    filtered = [
        v for v in all_values
        if any(term in _normalize_compact(v) for term in terms)
    ]

    print(f"   [tool] Filtered {table}[{column}] to {len(filtered)} matches "
          f"for terms {terms}")
    print(f"   [tool] Matches: {filtered}")

    matches = filtered if filtered else all_values
    # Cap to avoid blowing the token budget with huge lists
    if len(matches) > MAX_TOOL_RESPONSE_VALUES:
        print(f"   [tool] Capping to {MAX_TOOL_RESPONSE_VALUES} values.")
        matches = matches[:MAX_TOOL_RESPONSE_VALUES]
    return matches






# ──────────────────────────────────────────
# Token tracking
# ──────────────────────────────────────────
import csv
from datetime import datetime

TOKEN_LOG_PATH = os.environ.get("TOKEN_LOG_PATH", "token_usage.csv")

# Accumulated tokens for the current question (reset by begin_question / flush_tokens)
_current_tokens: Dict[str, int] = {"input": 0, "output": 0}
_current_question: str = ""


def begin_question(question: str) -> None:
    """Call before processing a new question to reset the per-question counter."""
    global _current_tokens, _current_question
    _current_tokens   = {"input": 0, "output": 0}
    _current_question = question


def _accumulate_tokens(response: Any) -> None:
    """Extract token counts from an OpenAI response and add to the running total."""
    try:
        usage = response.usage
        _current_tokens["input"]  += getattr(usage, "prompt_tokens", 0) or 0
        _current_tokens["output"] += getattr(usage, "completion_tokens", 0) or 0
    except Exception:
        pass  # never break the pipeline over logging


def flush_tokens(success: bool = True) -> Dict[str, int]:
    """
    Write the accumulated token counts for the current question to the CSV log.
    Returns the counts so the caller can display them.
    """
    snapshot = dict(_current_tokens)
    total    = snapshot["input"] + snapshot["output"]

    row = {
        "timestamp":      datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S"),
        "question":       _current_question,
        "input_tokens":   snapshot["input"],
        "output_tokens":  snapshot["output"],
        "total_tokens":   total,
        "model":          AZURE_OPENAI_DEPLOYMENT,
        "success":        success,
    }

    file_exists = os.path.isfile(TOKEN_LOG_PATH)
    try:
        with open(TOKEN_LOG_PATH, "a", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=row.keys())
            if not file_exists:
                writer.writeheader()
            writer.writerow(row)
        print(f"[tokens] input={snapshot['input']}  output={snapshot['output']}  "
              f"total={total}  → {TOKEN_LOG_PATH}")
    except Exception as exc:
        print(f"[tokens] Warning: could not write token log: {exc}")

    return snapshot






# ──────────────────────────────────────────
# Prompt engineering
# ──────────────────────────────────────────

DAX_SYSTEM_PROMPT = """
You are a Power BI DAX expert. Your job is to translate business questions into
correct DAX queries against a provided semantic model schema.

## Output format
Return valid JSON with exactly two keys:
  "dax"   : the full DAX query string
  "notes" : one sentence explaining what the query does

No markdown fences, no text outside the JSON object.

## How DAX evaluation works — read this before writing any query

Understanding DAX's execution model will help you avoid the most common errors.

DAX is not SQL. There are no "rows" in the traditional sense — instead, DAX
operates on filter contexts and row contexts propagated through relationships.
The key things to understand for writing correct EVALUATE queries are:

**Measures vs Columns**
The schema distinguishes between "columns" and "measures". They behave very
differently and must never be confused:

- A *column* (e.g. 'Sales'[Region]) is a physical value stored in a table.
  It can be used in groupBy arguments, FILTER predicates, and relationships.

- A *measure* (e.g. [Net Sales]) is a DAX expression that is evaluated lazily
  inside a filter context. Measures do not exist as column values — they are
  computed on demand. This is why you cannot reference a measure like a column.

The practical consequence: measures listed in the schema already exist in the
model. You call them directly as [Measure Name] inside SUMMARIZECOLUMNS or
ADDCOLUMNS. You never need to re-declare them with DEFINE MEASURE. In fact,
wrapping an existing measure in DEFINE MEASURE will break the query, because
inside the DEFINE block there is no filter context to evaluate [Measure Name]
against — DAX will report "cannot be determined".

DEFINE MEASURE is only appropriate when you need a *brand-new* calculated
expression that does not exist in the model (e.g. a custom ratio).

**The right way to use measures**

Use SUMMARIZECOLUMNS to group by columns and pull in measure values:

    EVALUATE
        SUMMARIZECOLUMNS(
            'DimCustomer'[CustomerName],
            "Net Sales", [Net Sales]
        )

To filter by a measure value, wrap SUMMARIZECOLUMNS in FILTER. Note that
inside the FILTER predicate you reference the column alias you defined ("Net Sales"),
not the measure directly:

    EVALUATE
        FILTER(
            SUMMARIZECOLUMNS(
                'DimCustomer'[CustomerName],
                "Net Sales", [Net Sales]
            ),
            [Net Sales] > 0
        )

## CRITICAL: where filters go in SUMMARIZECOLUMNS

A filter that constrains a column you are ALSO grouping by must be passed as a
top-level FILTER argument of SUMMARIZECOLUMNS — NEVER buried inside a CALCULATE
that wraps the measure.

If you put the filter inside CALCULATE instead of at the SUMMARIZECOLUMNS level,
the group-by column is left unconstrained in the query's own filter context, so
SUMMARIZECOLUMNS enumerates EVERY value of that column that exists in the
dimension (including stale/unused years like 1900, 1901, …), and the measure
returns the same identical total on every single row. This is a very common and
very wrong result.

WRONG — produces one row per year that exists in the calendar, all with the same total:

    EVALUATE
        SUMMARIZECOLUMNS(
            'W_CALENDAR_D'[Year],
            "Net Sales", CALCULATE([Net Sales], 'W_CALENDAR_D'[Year] = 2025)
        )

RIGHT — year filter is a top-level FILTER argument, so only 2025 is grouped:

    EVALUATE
        SUMMARIZECOLUMNS(
            'W_CALENDAR_D'[Year],
            FILTER('W_CALENDAR_D', 'W_CALENDAR_D'[Year] = 2025),
            "Net Sales", [Net Sales]
        )

## CRITICAL: single-scalar questions → use ROW(), not SUMMARIZECOLUMNS

If the question asks for ONE aggregate number with NO breakdown requested
(e.g. "how much X was sold in 2025?", "total revenue last quarter",
"X vs Y" comparisons), do NOT group by any column. Grouping by a column you
then filter to a single value is what produces the phantom-rows bug above.

Instead use ROW(), which evaluates each measure in its own CALCULATE context.
Time filters (year, quarter, month) and entity/class filters go INSIDE the
CALCULATE here — that is correct, because ROW() has no group-by column to leave
unconstrained:

    EVALUATE
        ROW(
            "Net Sales 2025", CALCULATE(
                [Net Sales],
                'W_CALENDAR_D'[Year] = 2025,
                FILTER('W_ICLASS_MBB0REP_D',
                       'W_ICLASS_MBB0REP_D'[ItemClassDesc] IN {
                           "Imported Rugs          07",
                           "Imported Rugs          10",
                           "Imported Rugs          31"
                       })
            )
        )

Multi-value single-row comparison ("Q1 2024 vs Q1 2026"):

    EVALUATE
        ROW(
            "Q1 2024 Sales", CALCULATE(
                [Net Sales],
                FILTER('W_CALENDAR_D', 'W_CALENDAR_D'[Year] = 2024 && 'W_CALENDAR_D'[Quarter] = 1),
                FILTER('W_CUST_MBBFREP_D', 'W_CUST_MBBFREP_D'[AlphaSortName] = "WAL-MART STORES")
            ),
            "Q1 2026 Sales", CALCULATE(
                [Net Sales],
                FILTER('W_CALENDAR_D', 'W_CALENDAR_D'[Year] = 2026 && 'W_CALENDAR_D'[Quarter] = 1),
                FILTER('W_CUST_MBBFREP_D', 'W_CUST_MBBFREP_D'[AlphaSortName] = "WAL-MART STORES")
            )
        )

Decision rule:
  - Result is a fixed set of named scalar values, no breakdown → ROW()
  - Result groups by one or more dimension columns → SUMMARIZECOLUMNS,
    with any filter on a grouped column passed as a top-level FILTER argument

## Matching values from get_column_values

When you call get_column_values you receive a list of candidate stored values.
How MANY you select depends on whether the user named a SPECIFIC ENTITY or a
CATEGORY / CLASS. These are different and must be handled differently:

- **Specific named entity** — a single customer, sales rep, store, or person
  (e.g. "Home Depot", "John Smith", "Walmart"):
  Select EVERY VALUE you see. At least a region or some specification is detailed on
  the user query. In that case, see which name match better that specification.

- **Category, class, or grouping dimension** — an item class, product category,
  region, segment, etc. (e.g. "imported rugs", "outdoor furniture", "rugs"):
  The user is naming a GROUP that may legitimately span MULTIPLE stored values.
  Examine ALL returned candidates and select EVERY value that genuinely belongs
  to that group. Exclude values that merely share a word but are a different
  class. Then filter using IN {...} listing all matching values.

  Example — user asks for "imported rugs", get_column_values returns:
      'CHENDI RUGS            10', 'Flemish Imported Rugs  31',
      'Imported Rugs          07', 'Imported Rugs          10',
      'Imported Rugs          31', 'Outdoor Rugs           10',
      'Printed Rugs           31', 'Specialty Rugs         07', ...
  Correct selection: the three "Imported Rugs" classes (07, 10, 31). Use
  judgment on borderline cases like "Flemish Imported Rugs" — include it only if
  it plausibly fits the user's intent. EXCLUDE unrelated classes that merely
  contain the word "Rugs" (Outdoor, Printed, Chendi, Specialty).

  Resulting filter:
      FILTER('W_ICLASS_MBB0REP_D',
             'W_ICLASS_MBB0REP_D'[ItemClassDesc] IN {
                 "Imported Rugs          07",
                 "Imported Rugs          10",
                 "Imported Rugs          31"
             })

- **Copy returned values VERBATIM** — including any trailing spaces, padding, or
  numeric codes (e.g. "Imported Rugs          31"). IN {...} and = perform
  EXACT string matching. Do not trim, re-pad, drop the trailing code, or
  reformat the value in any way. Use the strings exactly as get_column_values
  returned them.

**Year-over-year**
If a YoY comparison is requested with no explicit time period, compare
year-to-date (DATESYTD) for the current year vs the same period last year
using SAMEPERIODLASTYEAR or DATEADD.

**General advice**
Prefer SUMMARIZECOLUMNS over SUMMARIZE for queries that include measures — it
handles blank rows and cross-filter context more predictably. Only fall back to
ADDCOLUMNS(SUMMARIZE(...)) if you need to add computed columns to an existing
row set.

**Currency**
It's mandatory to group money and sales values by CurrencyID from W_SHPEXT_F. It's important because sales values has different currencies. 

**Gross sales**
When you are asked about gross sales, OrderType_Categorical (from W_SHPEXT_F) should be equal to Customer Orders. 
W_SHPEXT_F[OrderType_Categorical]=\"CO-Customer Orders\". Read "name": "Gross Sales".

Credit Memos are "negative sales". Adjustments, returns, cancellations, etc.


""".strip()


MAX_TOOL_RESPONSE_VALUES = 50   # cap on values returned per get_column_values call


def _prune_schema(schema: Dict[str, Any], question: str) -> Dict[str, Any]:
    """
    Return a leaner copy of the schema containing only tables/columns/measures
    that are plausibly relevant to the question.

    Strategy:
      - Always keep tables whose name matches a question keyword
      - Always keep all measures (they're small and any could be needed)
      - For tables that don't match by name, keep them only if at least one
        of their columns/measures matches a question keyword
      - Strip the 'expression' field from measures (saves tokens, Gemini doesn't need it)

    If no tables survive pruning (very abstract question), return the full schema
    so Gemini still has something to work with.
    """
    terms = set(_extract_search_terms(question))
    # Always keep calendar/date tables — almost every query needs them
    date_keywords = {"calendar", "date", "time", "period", "year", "month", "quarter"}

    pruned_tables = []
    for table in schema.get("tables", []):
        tname = _normalize_compact(table.get("name", ""))

        # Slim down measures — drop expression, keep name/description/format
        slim_measures = [
            {k: v for k, v in m.items() if k != "expression"}
            for m in table.get("measures", [])
        ]

        slim_table = {
            **{k: v for k, v in table.items() if k not in ("columns", "measures")},
            "columns":  table.get("columns", []),
            "measures": slim_measures,
        }

        # Always include if table name matches a term or is a date/calendar table
        if any(t in tname for t in terms) or any(d in tname for d in date_keywords):
            pruned_tables.append(slim_table)
            continue

        # Include if any column or measure name matches a term
        all_names = (
            [_normalize_compact(c.get("name", "")) for c in table.get("columns", [])]
            + [_normalize_compact(m.get("name", "")) for m in table.get("measures", [])]
        )
        if any(any(t in n for t in terms) for n in all_names if n):
            pruned_tables.append(slim_table)

    if not pruned_tables:
        # Fallback: return schema with just measure expressions stripped
        return {
            **schema,
            "tables": [
                {**t, "measures": [{k: v for k, v in m.items() if k != "expression"}
                                   for m in t.get("measures", [])]}
                for t in schema.get("tables", [])
            ],
        }

    return {**schema, "tables": pruned_tables}


def build_dax_instruction(schema: Dict[str, Any], question: str,
                          prior_dax: str = "", error_msg: str = "") -> str:
    prompt = DAX_SYSTEM_PROMPT
    # Prune schema to only relevant tables before serialising
    lean_schema = _prune_schema(schema, question)
    parts = [
        prompt,
        "\n\nSCHEMA:\n" + json.dumps(lean_schema, ensure_ascii=False, indent=2),
        "\n\nQUESTION:\n" + question,
    ]
    if prior_dax and error_msg:
        parts.append(
            f"\n\nPREVIOUS ATTEMPT FAILED — fix the query below.\n"
            f"FAILED DAX:\n{prior_dax}\n\n"
            f"ERROR FROM POWER BI API:\n{error_msg}\n\n"
            f"Common causes of '[Measure] cannot be determined' errors:\n"
            f"  - You used DEFINE MEASURE to wrap an existing model measure (never do this)\n"
            f"  - You referenced [Measure] inside FILTER() without a surrounding table context\n"
            f"  - Solution: use [Measure Name] directly inside SUMMARIZECOLUMNS, no DEFINE needed\n\n"
            f"Produce a corrected version. Return JSON only."
        )
    else:
        parts.append("\n\nReturn JSON only.")
    return "".join(parts)


ANSWER_SYSTEM_PROMPT = """
You are a business analyst presenting query results to an executive.
The data rows below were returned directly from a live database query that already
applied all the correct filters. Trust the data completely — if rows are present,
the data exists. Never say information is missing, unavailable, or not found.
Simply read the values from the rows and present them clearly.
One direct sentence first, then bullet points if there are multiple items.
Format numbers with commas. Do not invent values not in the data.
""".strip()


def build_answer_instruction(question: str, df: pd.DataFrame) -> str:
    sample = df.head(MAX_ROWS_TO_LLM).to_dict(orient="records")
    columns_hint = ", ".join(f'"{c}"' for c in df.columns)
    return (
        ANSWER_SYSTEM_PROMPT
        + f"\n\nQ: {question}"
        + f"\n\nColumns in the result: {columns_hint}"
        + f"\n\nDATA ({len(sample)} row(s) — all values are real and correct):\n"
        + json.dumps(sample, ensure_ascii=False, default=str, separators=(",", ":"))
        + "\n\nA:"
    )


def parse_json_strict(text: str) -> Dict[str, Any]:
    text = text.strip()
    # Strip markdown code fences if the model added them despite instructions
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", text, flags=re.DOTALL)
        if not m:
            raise ValueError(f"No JSON object found in model output:\n{text[:300]}")
        return json.loads(m.group(0))


# ──────────────────────────────────────────
# Pipeline: question → DAX → data → answer
# ──────────────────────────────────────────

DAX_MAX_RETRIES = 3  # how many times to ask Gemini to fix a broken DAX


def _extract_pbi_error(exc: RuntimeError) -> str:
    """Pull the human-readable detail message out of a Power BI 400 error."""
    text = str(exc)
    # Try to parse the JSON body embedded in the RuntimeError message
    m = re.search(r"\{.*\}", text, re.DOTALL)
    if not m:
        return text
    try:
        body = json.loads(m.group(0))
        details = (
            body.get("error", {})
                .get("pbi.error", {})
                .get("details", [])
        )
        messages = [d["detail"]["value"] for d in details if "detail" in d]
        return "\n".join(messages) if messages else text
    except Exception:
        return text


def ask(question: str, schema: Dict[str, Any], token: str) -> str:
    """
    Full pipeline for one question. Returns the natural-language answer.
    Token usage is accumulated across all Gemini calls and flushed to CSV at the end.
    """
    begin_question(question)
    prior_dax = ""
    error_msg = ""
    success   = False

    try:
        for attempt in range(1, DAX_MAX_RETRIES + 1):

            # ── Step 1: Generate (or fix) DAX ────────────────────────────────────
            if attempt == 1:
                print("\n[1/3] Generating DAX query…")
                raw = run_function_calling_turn(question, schema, token)
            else:
                print(f"\n[1/3] Fixing DAX query (attempt {attempt}/{DAX_MAX_RETRIES})…")
                raw = generate_response(
                    build_dax_instruction(schema, question, prior_dax, error_msg)
                )

            try:
                obj = parse_json_strict(raw)
            except (ValueError, json.JSONDecodeError) as exc:
                return (
                    f"❌ Could not parse model output as JSON: {exc}"
                    f"\n\nRaw output:\n{raw[:500]}"
                )

            dax   = obj.get("dax", "").strip()
            notes = obj.get("notes", "")

            if "EVALUATE" not in dax.upper():
                return f"❌ Generated DAX contains no EVALUATE statement:\n{dax[:300]}"

            print(f"   Notes: {notes}")
            print(f"   DAX:\n{textwrap.indent(dax, '   ')}")

            # ── Step 2: Execute DAX ───────────────────────────────────────────────
            print("\n[2/3] Executing DAX on Power BI…")
            try:
                result_json = execute_dax(token, dax)
                break
            except RuntimeError as exc:
                error_msg = _extract_pbi_error(exc)
                prior_dax = dax
                print(f"   ⚠️  Power BI error (attempt {attempt}): {error_msg}")
                if attempt == DAX_MAX_RETRIES:
                    return (
                        f"❌ DAX failed after {DAX_MAX_RETRIES} attempts.\n"
                        f"Last error: {error_msg}"
                    )
                continue

        df = result_to_dataframe(result_json)
        if df.empty:
            return "ℹ️ The query returned no rows."

        print(f"   Returned {len(df)} row(s), {len(df.columns)} column(s).")

        # ── Step 3: Summarise ─────────────────────────────────────────────────────
        print("\n[3/3] Summarising results…")
        answer  = generate_response(build_answer_instruction(question, df))
        success = True
        return answer

    finally:
        flush_tokens(success=success)


# ──────────────────────────────────────────
# CLI entry point
# ──────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Power BI Chatbot (Gemini + DAX)")
    parser.add_argument("--once", metavar="QUESTION", help="Ask one question and exit")
    parser.add_argument(
        "--auth",
        choices=["auto", "env", "powershell", "devicecode"],
        default="auto",
        help="Authentication method (default: auto)",
    )
    args = parser.parse_args()

    # Load schema
    with open(SCHEMA_PATH, "r", encoding="utf-8") as f:
        schema = json.load(f)

    # Authenticate once
    token = acquire_token(args.auth)
    print("[auth] Token acquired ✓\n")

    if args.once:
        print(f"QUESTION: {args.once}")
        print("\nANSWER:")
        print(ask(args.once, schema, token))
        return

    # Interactive chat loop
    print("Power BI Chatbot — type 'exit' to quit.\n")
    while True:
        try:
            question = input("You: ").strip()
        except (EOFError, KeyboardInterrupt):
            break

        if not question:
            continue
        if question.lower() in {"exit", "quit", "q"}:
            break

        answer = ask(question, schema, token)
        print(f"\nBot: {answer}\n")


if __name__ == "__main__":
    main()
