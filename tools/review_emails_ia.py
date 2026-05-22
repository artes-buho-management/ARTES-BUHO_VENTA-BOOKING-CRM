import argparse
import json
import os
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import asdict, dataclass
from typing import Dict, List, Optional, Set, Tuple

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


DEFAULT_SPREADSHEET_ID = "REPLACE_WITH_SHEET_ID"
DEFAULT_KEY_FILE = r"C:\Users\elrub\Desktop\CARPETA CODEX\secrets\robot-codex-key-20260308-220232.json"

IA_HEADER = "CORREO REVISADO"
MERGE_HEADER = "Merge status"
VALIDATION_VALUES = ["BIEN", "MAL", "CORREGIDO"]
STATUS_GOOD = "BIEN"
STATUS_BAD = "MAL"
STATUS_FIXED = "CORREGIDO"

LEGACY_STATUS_MAP = {
    "REVISADO": STATUS_GOOD,
    "NO": STATUS_BAD,
    "PTE": STATUS_BAD,
    "ERROR": STATUS_BAD,
    "": STATUS_BAD,
}

EMAIL_RE = re.compile(r"^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+$")
EMAIL_CAPTURE_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
SPLIT_RE = re.compile(r"[\s,;|/]+")
PLACEHOLDER_TOKENS = {
    "",
    "no disponible",
    "sin email",
    "n/a",
    "na",
    "-",
    "pendiente",
    "no tiene",
    "none",
    "no email",
}
KNOWN_DOMAIN_TYPOS = {
    "gamil.com",
    "gmai.com",
    "hotnail.com",
    "hotmal.com",
    "outlok.com",
    "yaho.com",
}

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

WRITE_MIN_INTERVAL_SECONDS = 1.05
_LAST_WRITE_TS = 0.0


def execute_request(req, is_write: bool = False, max_retries: int = 8):
    global _LAST_WRITE_TS
    attempt = 0
    while True:
        try:
            if is_write:
                now = time.time()
                wait = WRITE_MIN_INTERVAL_SECONDS - (now - _LAST_WRITE_TS)
                if wait > 0:
                    time.sleep(wait)
            result = req.execute()
            if is_write:
                _LAST_WRITE_TS = time.time()
            return result
        except HttpError as err:
            status = getattr(err.resp, "status", None)
            is_rate_limit = status in (429, 503)
            if status == 403:
                content = ""
                try:
                    content = err.content.decode("utf-8", errors="ignore").lower()
                except Exception:
                    content = str(err).lower()
                is_rate_limit = "rate limit" in content or "quota" in content

            if is_rate_limit and attempt < max_retries:
                sleep_s = min(60.0, 2 ** attempt + 0.5)
                time.sleep(sleep_s)
                attempt += 1
                continue
            raise


@dataclass
class SheetRunStats:
    title: str
    email_col: int
    call_col: int
    ia_col: int
    merge_col: int
    rows_data: int
    processed: int
    changed: int
    good: int
    bad: int
    fixed: int
    corrected: int
    web_lookups: int


def col_to_letter(col: int) -> str:
    result = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result
    return result


def normalize_text(value: str) -> str:
    return (value or "").strip().lower()


def normalize_status(value: str) -> str:
    raw = (value or "").strip().upper()
    if raw in VALIDATION_VALUES:
        return raw
    return LEGACY_STATUS_MAP.get(raw, STATUS_BAD)


def extract_email(raw: str) -> Optional[str]:
    text = (raw or "").strip()
    if not text:
        return None
    low = normalize_text(text)
    if low in PLACEHOLDER_TOKENS:
        return None
    if "@" in text and " " not in text and "," not in text and ";" not in text and "|" not in text:
        return text.strip().strip(".").lower()
    for token in SPLIT_RE.split(text):
        token = token.strip().strip(".").lower()
        if "@" in token:
            return token
    return None


def local_email_check(email: str) -> Tuple[bool, str]:
    if not email or "@" not in email:
        return False, "missing_at"
    if email.count("@") != 1:
        return False, "multiple_at"
    if not EMAIL_RE.match(email):
        return False, "regex_fail"
    local, domain = email.split("@", 1)
    if ".." in local or ".." in domain:
        return False, "double_dot"
    if domain in KNOWN_DOMAIN_TYPOS:
        return False, "known_typo_domain"
    if domain.startswith("-") or domain.endswith("-"):
        return False, "bad_domain_dash"
    if len(local) > 64 or len(email) > 254:
        return False, "length"
    tld = domain.split(".")[-1]
    if len(tld) < 2:
        return False, "short_tld"
    return True, "syntax_ok"


def domain_has_dns(domain: str, cache: Dict[str, Optional[bool]]) -> Optional[bool]:
    if domain in cache:
        return cache[domain]

    def query_dns(record_type: str) -> Optional[bool]:
        url = "https://dns.google/resolve?" + urllib.parse.urlencode({"name": domain, "type": record_type})
        req = urllib.request.Request(url, headers={"Accept": "application/json", "User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=8) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
        if payload.get("Status") == 0 and payload.get("Answer"):
            return True
        return False

    try:
        if query_dns("MX"):
            cache[domain] = True
            return True
        if query_dns("A"):
            cache[domain] = True
            return True
        cache[domain] = False
        return False
    except (urllib.error.URLError, TimeoutError, OSError, ValueError):
        cache[domain] = None
        return None


def is_email_header(header: str) -> bool:
    h = normalize_text(header)
    return "email" in h or "e-mail" in h


def is_call_header(header: str) -> bool:
    h = normalize_text(header)
    return "llamada" in h


def quote_sheet(title: str) -> str:
    return "'" + title.replace("'", "''") + "'"


def get_headers(values_api, spreadsheet_id: str, sheet_title: str) -> List[str]:
    rng = f"{quote_sheet(sheet_title)}!1:1"
    data = execute_request(values_api.get(spreadsheetId=spreadsheet_id, range=rng))
    row = data.get("values", [[]])[0] if data.get("values") else []
    return [str(x).strip() for x in row]


def find_col(headers: List[str], matcher) -> int:
    for idx, header in enumerate(headers, start=1):
        if matcher(header):
            return idx
    return 0


def find_ia_col(headers: List[str]) -> int:
    for idx, header in enumerate(headers, start=1):
        h = normalize_text(header)
        if h == normalize_text(IA_HEADER) or h == "correo revisado ia":
            return idx
    return 0


def find_merge_col(headers: List[str]) -> int:
    target = normalize_text(MERGE_HEADER)
    for idx, header in enumerate(headers, start=1):
        if normalize_text(header) == target:
            return idx
    return 0


def find_context_cols(headers: List[str]) -> Dict[str, int]:
    name_col = 0
    municipio_col = 0
    provincia_col = 0

    for idx, h in enumerate(headers, start=1):
        s = normalize_text(h)
        if name_col == 0 and "nombre" in s and "contacto" not in s:
            name_col = idx
        if municipio_col == 0 and ("municipio" in s or "poblacion" in s):
            municipio_col = idx
        if provincia_col == 0 and "provincia" in s:
            provincia_col = idx

    return {
        "name_col": name_col,
        "municipio_col": municipio_col,
        "provincia_col": provincia_col,
    }


def build_color_rule(formula: str, ranges: List[dict], rgb: Tuple[float, float, float]) -> dict:
    return {
        "ranges": ranges,
        "booleanRule": {
            "condition": {
                "type": "CUSTOM_FORMULA",
                "values": [{"userEnteredValue": formula}],
            },
            "format": {
                "backgroundColor": {
                    "red": rgb[0],
                    "green": rgb[1],
                    "blue": rgb[2],
                }
            },
        },
    }


def apply_status_color_rules(
    sheets_api,
    spreadsheet_id: str,
    sheet_meta: dict,
    ia_col: int,
    row_count: int,
    col_count: int,
    dry_run: bool,
):
    sheet_id = sheet_meta["properties"]["sheetId"]
    existing_rules = sheet_meta.get("conditionalFormats", []) or []

    col_letter = col_to_letter(ia_col)
    formula_good = f'=${col_letter}2="{STATUS_GOOD}"'
    formula_bad = f'=${col_letter}2="{STATUS_BAD}"'
    formula_fixed = f'=${col_letter}2="{STATUS_FIXED}"'

    codex_formulas = {
        formula_good,
        formula_bad,
        formula_fixed,
        f'=${col_letter}2="REVISADO"',
        f'=${col_letter}2="NO"',
        f'=${col_letter}2="ERROR"',
        f'=${col_letter}2="PTE"',
    }

    delete_indexes: List[int] = []
    for idx, rule in enumerate(existing_rules):
        cond = ((rule or {}).get("booleanRule") or {}).get("condition") or {}
        if cond.get("type") != "CUSTOM_FORMULA":
            continue
        values = cond.get("values") or []
        if not values:
            continue
        formula = (values[0] or {}).get("userEnteredValue")
        if formula in codex_formulas:
            delete_indexes.append(idx)

    requests: List[dict] = []
    for idx in sorted(delete_indexes, reverse=True):
        requests.append(
            {
                "deleteConditionalFormatRule": {
                    "sheetId": sheet_id,
                    "index": idx,
                }
            }
        )

    remaining = len(existing_rules) - len(delete_indexes)
    row_range = {
        "sheetId": sheet_id,
        "startRowIndex": 1,
        "endRowIndex": row_count,
        "startColumnIndex": 0,
        "endColumnIndex": max(col_count, ia_col),
    }

    new_rules = [
        build_color_rule(formula_good, [row_range], (0.85, 0.92, 0.83)),
        build_color_rule(formula_bad, [row_range], (0.96, 0.80, 0.80)),
        build_color_rule(formula_fixed, [row_range], (0.85, 0.89, 0.95)),
    ]

    for i, rule in enumerate(new_rules):
        requests.append(
            {
                "addConditionalFormatRule": {
                    "index": remaining + i,
                    "rule": rule,
                }
            }
        )

    if not dry_run and requests:
        execute_request(
            sheets_api.batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests}),
            is_write=True,
        )


def ensure_review_column(
    sheets_api,
    values_api,
    spreadsheet_id: str,
    sheet_meta: dict,
    dry_run: bool,
) -> dict:
    props = sheet_meta["properties"]
    title = props["title"]
    sheet_id = props["sheetId"]
    row_count = int(props.get("gridProperties", {}).get("rowCount", 1000))
    col_count = int(props.get("gridProperties", {}).get("columnCount", 26))

    headers = get_headers(values_api, spreadsheet_id, title)
    email_col = find_col(headers, is_email_header)
    call_col = find_col(headers, is_call_header)
    ia_col = find_ia_col(headers)
    merge_col = find_merge_col(headers)

    if email_col == 0:
        return {
            "title": title,
            "email_col": 0,
            "call_col": 0,
            "ia_col": 0,
            "merge_col": 0,
            "row_count": row_count,
            "col_count": col_count,
            "headers": headers,
            "context": find_context_cols(headers),
        }

    ia_created = False
    if ia_col == 0:
        insert_after = call_col if call_col > 0 else email_col
        insert_req = {
            "insertDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": insert_after,
                    "endIndex": insert_after + 1,
                },
                "inheritFromBefore": True,
            }
        }
        if not dry_run:
            execute_request(
                sheets_api.batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": [insert_req]}),
                is_write=True,
            )
        ia_col = insert_after + 1
        ia_created = True
        col_count += 1
        if not dry_run:
            header_range = f"{quote_sheet(title)}!{col_to_letter(ia_col)}1"
            execute_request(
                values_api.update(
                    spreadsheetId=spreadsheet_id,
                    range=header_range,
                    valueInputOption="USER_ENTERED",
                    body={"values": [[IA_HEADER]]},
                ),
                is_write=True,
            )
        if ia_col - 1 <= len(headers):
            headers.insert(ia_col - 1, IA_HEADER)
        else:
            while len(headers) < ia_col - 1:
                headers.append("")
            headers.append(IA_HEADER)

    current_header = headers[ia_col - 1] if ia_col - 1 < len(headers) else ""
    if normalize_text(current_header) != normalize_text(IA_HEADER):
        if not dry_run:
            header_range = f"{quote_sheet(title)}!{col_to_letter(ia_col)}1"
            execute_request(
                values_api.update(
                    spreadsheetId=spreadsheet_id,
                    range=header_range,
                    valueInputOption="USER_ENTERED",
                    body={"values": [[IA_HEADER]]},
                ),
                is_write=True,
            )
        if ia_col - 1 < len(headers):
            headers[ia_col - 1] = IA_HEADER

    if merge_col == 0:
        insert_after = ia_col
        insert_req = {
            "insertDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": insert_after,
                    "endIndex": insert_after + 1,
                },
                "inheritFromBefore": True,
            }
        }
        if not dry_run:
            execute_request(
                sheets_api.batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": [insert_req]}),
                is_write=True,
            )
        merge_col = insert_after + 1
        col_count += 1
        if not dry_run:
            header_range = f"{quote_sheet(title)}!{col_to_letter(merge_col)}1"
            execute_request(
                values_api.update(
                    spreadsheetId=spreadsheet_id,
                    range=header_range,
                    valueInputOption="USER_ENTERED",
                    body={"values": [[MERGE_HEADER]]},
                ),
                is_write=True,
            )
        if merge_col - 1 <= len(headers):
            headers.insert(merge_col - 1, MERGE_HEADER)
        else:
            while len(headers) < merge_col - 1:
                headers.append("")
            headers.append(MERGE_HEADER)

    current_merge_header = headers[merge_col - 1] if merge_col - 1 < len(headers) else ""
    if normalize_text(current_merge_header) != normalize_text(MERGE_HEADER):
        if not dry_run:
            header_range = f"{quote_sheet(title)}!{col_to_letter(merge_col)}1"
            execute_request(
                values_api.update(
                    spreadsheetId=spreadsheet_id,
                    range=header_range,
                    valueInputOption="USER_ENTERED",
                    body={"values": [[MERGE_HEADER]]},
                ),
                is_write=True,
            )
        if merge_col - 1 < len(headers):
            headers[merge_col - 1] = MERGE_HEADER

    # Keep "Merge status" as the last column to protect YAMM strategy workflows.
    if merge_col and merge_col != col_count:
        old_merge_col = merge_col
        if not dry_run:
            requests = [
                {
                    "insertDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": col_count,
                            "endIndex": col_count + 1,
                        },
                        "inheritFromBefore": True,
                    }
                },
                {
                    "copyPaste": {
                        "source": {
                            "sheetId": sheet_id,
                            "startRowIndex": 0,
                            "endRowIndex": row_count,
                            "startColumnIndex": merge_col - 1,
                            "endColumnIndex": merge_col,
                        },
                        "destination": {
                            "sheetId": sheet_id,
                            "startRowIndex": 0,
                            "endRowIndex": row_count,
                            "startColumnIndex": col_count,
                            "endColumnIndex": col_count + 1,
                        },
                        "pasteType": "PASTE_NORMAL",
                        "pasteOrientation": "NORMAL",
                    }
                },
                {
                    "deleteDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": merge_col - 1,
                            "endIndex": merge_col,
                        }
                    }
                },
            ]
            execute_request(
                sheets_api.batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests}),
                is_write=True,
            )
        # After moving, merge returns to previous last used index (1-based).
        merge_col = col_count
        if old_merge_col - 1 < len(headers):
            value = headers.pop(old_merge_col - 1)
            headers.append(value)

    # Always enforce current dropdown values.
    if not dry_run:
        validation_req = {
            "setDataValidation": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 1,
                    "endRowIndex": row_count,
                    "startColumnIndex": ia_col - 1,
                    "endColumnIndex": ia_col,
                },
                "rule": {
                    "condition": {
                        "type": "ONE_OF_LIST",
                        "values": [{"userEnteredValue": v} for v in VALIDATION_VALUES],
                    },
                    "inputMessage": "Estado de revision del email",
                    "strict": True,
                    "showCustomUi": True,
                },
            }
        }
        execute_request(
            sheets_api.batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": [validation_req]}),
            is_write=True,
        )

    if ia_created:
        # Metadata from initial read doesn't include new CF state; start clean for this pass.
        sheet_meta = {
            "properties": props,
            "conditionalFormats": sheet_meta.get("conditionalFormats", []),
        }

    apply_status_color_rules(
        sheets_api=sheets_api,
        spreadsheet_id=spreadsheet_id,
        sheet_meta=sheet_meta,
        ia_col=ia_col,
        row_count=row_count,
        col_count=col_count,
        dry_run=dry_run,
    )

    context = find_context_cols(headers)
    return {
        "title": title,
        "email_col": email_col,
        "call_col": call_col,
        "ia_col": ia_col,
        "merge_col": merge_col,
        "row_count": row_count,
        "col_count": col_count,
        "headers": headers,
        "context": context,
    }


def fetch_column_values(values_api, spreadsheet_id: str, title: str, col: int) -> List[str]:
    rng = f"{quote_sheet(title)}!{col_to_letter(col)}:{col_to_letter(col)}"
    data = execute_request(values_api.get(spreadsheetId=spreadsheet_id, range=rng))
    rows = data.get("values", [])
    out: List[str] = []
    for row in rows:
        out.append(str(row[0]) if row else "")
    return out


def fetch_multiple_columns(values_api, spreadsheet_id: str, title: str, cols: Set[int]) -> Dict[int, List[str]]:
    out: Dict[int, List[str]] = {}
    for col in sorted([c for c in cols if c > 0]):
        out[col] = fetch_column_values(values_api, spreadsheet_id, title, col)
    return out


def pick_best_candidate(candidates: List[str], current_email: Optional[str]) -> Optional[str]:
    if not candidates:
        return None

    lowered = [c.lower() for c in candidates]
    if current_email and current_email.lower() in lowered:
        return current_email.lower()

    current_domain = ""
    if current_email and "@" in current_email:
        current_domain = current_email.split("@", 1)[1].lower()

    if current_domain:
        for cand in candidates:
            if cand.lower().endswith("@" + current_domain):
                return cand.lower()

    # Prefer shorter addresses if multiple were found.
    ordered = sorted(candidates, key=lambda x: (len(x), x))
    return ordered[0].lower()


def build_contact_query(name: str, municipio: str, provincia: str) -> str:
    parts = []
    for v in (name, municipio, provincia):
        vv = (v or "").strip()
        if vv:
            parts.append(vv)

    if not parts:
        return ""

    parts.append("email contacto")
    return " ".join(parts)


def fetch_text_url(url: str, timeout_s: int = 10) -> str:
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0",
            "Accept": "text/html,application/xhtml+xml",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout_s) as resp:
        return resp.read().decode("utf-8", errors="ignore")


def extract_emails_from_text(text: str) -> List[str]:
    found = set()
    for m in EMAIL_CAPTURE_RE.findall(text or ""):
        email = m.strip().strip(".").lower()
        ok, _ = local_email_check(email)
        if ok:
            found.add(email)
    return sorted(found)


def extract_duck_links(html: str) -> List[str]:
    links: List[str] = []
    if not html:
        return links

    for encoded in re.findall(r"uddg=([^&\"']+)", html):
        try:
            decoded = urllib.parse.unquote(encoded)
        except Exception:
            continue
        if decoded.startswith("http://") or decoded.startswith("https://"):
            links.append(decoded)

    for href in re.findall(r'href="(https?://[^\"]+)"', html):
        links.append(href)

    unique = []
    seen = set()
    for u in links:
        u2 = u.strip()
        if not u2 or u2 in seen:
            continue
        seen.add(u2)
        unique.append(u2)
    return unique


def search_emails_duck(
    query: str,
    query_cache: Dict[str, List[str]],
    page_cache: Dict[str, List[str]],
    max_links: int,
) -> List[str]:
    if query in query_cache:
        return query_cache[query]

    emails: Set[str] = set()

    try:
        search_url = "https://duckduckgo.com/html/?" + urllib.parse.urlencode({"q": query})
        html = fetch_text_url(search_url, timeout_s=10)
        for email in extract_emails_from_text(html):
            emails.add(email)

        if max_links > 0:
            links = extract_duck_links(html)
            for link in links[:max_links]:
                if link in page_cache:
                    for e in page_cache[link]:
                        emails.add(e)
                    continue
                try:
                    page_html = fetch_text_url(link, timeout_s=8)
                    page_emails = extract_emails_from_text(page_html)
                except Exception:
                    page_emails = []
                page_cache[link] = page_emails
                for e in page_emails:
                    emails.add(e)
    except Exception:
        pass

    out = sorted(emails)
    query_cache[query] = out
    return out


def email_is_valid(email: Optional[str], dns_cache: Dict[str, Optional[bool]], dns_mode: str) -> bool:
    if not email:
        return False
    ok, _ = local_email_check(email)
    if not ok:
        return False
    if dns_mode == "none":
        return True
    domain = email.split("@", 1)[1]
    dns = domain_has_dns(domain, dns_cache)
    return dns is not False


def batch_update_single_cells(values_api, spreadsheet_id: str, updates: List[Tuple[str, str]]):
    if not updates:
        return

    chunk_size = 150
    for i in range(0, len(updates), chunk_size):
        chunk = updates[i : i + chunk_size]
        data = [{"range": rg, "values": [[val]]} for rg, val in chunk]
        execute_request(
            values_api.batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={
                    "valueInputOption": "USER_ENTERED",
                    "data": data,
                },
            ),
            is_write=True,
        )


def review_sheet(
    values_api,
    spreadsheet_id: str,
    cfg: dict,
    dns_cache: Dict[str, Optional[bool]],
    web_query_cache: Dict[str, List[str]],
    web_page_cache: Dict[str, List[str]],
    dns_mode: str,
    web_mode: str,
    max_web_lookups: int,
    max_links: int,
    dry_run: bool,
) -> SheetRunStats:
    title = cfg["title"]
    email_col = cfg["email_col"]
    ia_col = cfg["ia_col"]
    call_col = cfg["call_col"]
    merge_col = cfg.get("merge_col", 0)

    context = cfg["context"]
    needed_cols = {
        email_col,
        ia_col,
        context.get("name_col", 0),
        context.get("municipio_col", 0),
        context.get("provincia_col", 0),
    }
    col_data = fetch_multiple_columns(values_api, spreadsheet_id, title, needed_cols)

    email_col_values = col_data.get(email_col, [])
    ia_col_values = col_data.get(ia_col, [])

    if len(email_col_values) <= 1:
        return SheetRunStats(
            title=title,
            email_col=email_col,
            call_col=call_col,
            ia_col=ia_col,
            merge_col=merge_col,
            rows_data=0,
            processed=0,
            changed=0,
            good=0,
            bad=0,
            fixed=0,
            corrected=0,
            web_lookups=0,
        )

    rows_data = len(email_col_values) - 1
    email_rows = email_col_values[1:]

    current_status_rows = ia_col_values[1:] if len(ia_col_values) > 1 else []
    if len(current_status_rows) < rows_data:
        current_status_rows += [""] * (rows_data - len(current_status_rows))

    name_values = col_data.get(context.get("name_col", 0), [""])
    municipio_values = col_data.get(context.get("municipio_col", 0), [""])
    provincia_values = col_data.get(context.get("provincia_col", 0), [""])

    def safe_row_value(arr: List[str], row_idx_1based: int) -> str:
        if not arr:
            return ""
        if row_idx_1based < len(arr):
            return arr[row_idx_1based] or ""
        return ""

    new_status_rows: List[str] = []
    email_corrections: List[Tuple[str, str]] = []

    processed = 0
    changed = 0
    good = 0
    bad = 0
    fixed = 0
    corrected = 0
    web_lookups_used = 0
    web_remaining = max(0, max_web_lookups)

    for idx in range(rows_data):
        row_number = idx + 2
        raw_status = (current_status_rows[idx] or "").strip().upper()
        cur_status = normalize_status(raw_status)

        raw_email = email_rows[idx]
        email = extract_email(raw_email)
        valid_email = email_is_valid(email, dns_cache, dns_mode)

        new_status = cur_status
        corrected_email: Optional[str] = None

        if cur_status == STATUS_FIXED and valid_email:
            new_status = STATUS_FIXED
        elif valid_email:
            new_status = STATUS_GOOD
        else:
            candidate = None
            if web_mode == "duck" and web_remaining > 0:
                name = safe_row_value(name_values, row_number)
                municipio = safe_row_value(municipio_values, row_number)
                provincia = safe_row_value(provincia_values, row_number)
                query = build_contact_query(name, municipio, provincia)
                if query:
                    web_remaining -= 1
                    web_lookups_used += 1
                    candidates = search_emails_duck(
                        query=query,
                        query_cache=web_query_cache,
                        page_cache=web_page_cache,
                        max_links=max_links,
                    )
                    candidate = pick_best_candidate(candidates, email)

            if candidate and email_is_valid(candidate, dns_cache, dns_mode):
                if email and candidate.lower() == email.lower():
                    new_status = STATUS_GOOD
                else:
                    new_status = STATUS_FIXED
                    corrected_email = candidate.lower()
            else:
                new_status = STATUS_BAD

        if new_status == STATUS_GOOD:
            good += 1
        elif new_status == STATUS_FIXED:
            fixed += 1
        else:
            bad += 1

        if new_status != raw_status:
            changed += 1

        if corrected_email:
            email_a1 = f"{quote_sheet(title)}!{col_to_letter(email_col)}{row_number}"
            email_corrections.append((email_a1, corrected_email))
            corrected += 1
            processed += 1
        elif new_status != cur_status or raw_status != new_status:
            processed += 1

        new_status_rows.append(new_status)

    if not dry_run and changed > 0:
        start = 2
        end = len(new_status_rows) + 1
        target_range = f"{quote_sheet(title)}!{col_to_letter(ia_col)}{start}:{col_to_letter(ia_col)}{end}"
        body = {"values": [[v] for v in new_status_rows]}
        execute_request(
            values_api.update(
                spreadsheetId=spreadsheet_id,
                range=target_range,
                valueInputOption="USER_ENTERED",
                body=body,
            ),
            is_write=True,
        )

    if not dry_run and email_corrections:
        batch_update_single_cells(values_api, spreadsheet_id, email_corrections)

    return SheetRunStats(
        title=title,
        email_col=email_col,
        call_col=call_col,
        ia_col=ia_col,
        merge_col=merge_col,
        rows_data=rows_data,
        processed=processed,
        changed=changed,
        good=good,
        bad=bad,
        fixed=fixed,
        corrected=corrected,
        web_lookups=web_lookups_used,
    )


def main():
    parser = argparse.ArgumentParser(description="Revision de emails CRM VENTA-BOOKING")
    parser.add_argument("--spreadsheet-id", default=DEFAULT_SPREADSHEET_ID)
    parser.add_argument("--key-file", default=DEFAULT_KEY_FILE)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--dns-mode", choices=["none", "mx_a"], default="none")
    parser.add_argument("--web-mode", choices=["off", "duck"], default="duck")
    parser.add_argument("--max-web-lookups", type=int, default=2)
    parser.add_argument("--max-web-links", type=int, default=0)
    parser.add_argument(
        "--report-file",
        default=r"C:\Users\elrub\Desktop\CARPETA CODEX\04_TEMPORAL\email_review_report.json",
    )
    args = parser.parse_args()

    if not os.path.exists(args.key_file):
        raise FileNotFoundError(f"No existe key file: {args.key_file}")

    creds = service_account.Credentials.from_service_account_file(args.key_file, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds, cache_discovery=False)
    sheets_api = service.spreadsheets()
    values_api = sheets_api.values()

    meta = execute_request(
        sheets_api.get(
            spreadsheetId=args.spreadsheet_id,
            fields=(
                "spreadsheetId,properties.title,"
                "sheets(properties(sheetId,title,hidden,gridProperties(rowCount,columnCount)),conditionalFormats)"
            ),
        )
    )

    dns_cache: Dict[str, Optional[bool]] = {}
    web_query_cache: Dict[str, List[str]] = {}
    web_page_cache: Dict[str, List[str]] = {}
    started = time.time()

    all_stats: List[SheetRunStats] = []
    for sheet in meta.get("sheets", []):
        props = sheet.get("properties", {})
        title = props.get("title", "")
        if props.get("hidden"):
            continue

        cfg = ensure_review_column(
            sheets_api=sheets_api,
            values_api=values_api,
            spreadsheet_id=args.spreadsheet_id,
            sheet_meta=sheet,
            dry_run=args.dry_run,
        )
        if cfg["email_col"] == 0 or cfg["ia_col"] == 0:
            continue

        stats = review_sheet(
            values_api=values_api,
            spreadsheet_id=args.spreadsheet_id,
            cfg=cfg,
            dns_cache=dns_cache,
            web_query_cache=web_query_cache,
            web_page_cache=web_page_cache,
            dns_mode=args.dns_mode,
            web_mode=args.web_mode,
            max_web_lookups=args.max_web_lookups,
            max_links=args.max_web_links,
            dry_run=args.dry_run,
        )
        all_stats.append(stats)

        print(
            f"[{title}] rows={stats.rows_data} processed={stats.processed} changed={stats.changed} "
            f"bien={stats.good} mal={stats.bad} corregido={stats.fixed} "
            f"corr_emails={stats.corrected} web={stats.web_lookups} "
            f"cols(email={stats.email_col},llamada={stats.call_col},rev={stats.ia_col},merge={stats.merge_col})"
        )

    total = {
        "spreadsheetId": args.spreadsheet_id,
        "title": meta.get("properties", {}).get("title", ""),
        "dryRun": args.dry_run,
        "dnsMode": args.dns_mode,
        "webMode": args.web_mode,
        "maxWebLookups": args.max_web_lookups,
        "generatedAt": time.strftime("%Y-%m-%d %H:%M:%S"),
        "seconds": round(time.time() - started, 2),
        "sheets": [asdict(s) for s in all_stats],
        "summary": {
            "sheetsProcessed": len(all_stats),
            "rowsData": sum(s.rows_data for s in all_stats),
            "processed": sum(s.processed for s in all_stats),
            "changed": sum(s.changed for s in all_stats),
            "bien": sum(s.good for s in all_stats),
            "mal": sum(s.bad for s in all_stats),
            "corregido": sum(s.fixed for s in all_stats),
            "emailsCorregidos": sum(s.corrected for s in all_stats),
            "webLookups": sum(s.web_lookups for s in all_stats),
            "uniqueDomainsChecked": len(dns_cache),
            "queryCache": len(web_query_cache),
        },
    }

    os.makedirs(os.path.dirname(args.report_file), exist_ok=True)
    with open(args.report_file, "w", encoding="utf-8") as f:
        json.dump(total, f, ensure_ascii=False, indent=2)

    print("REPORT:", args.report_file)
    print("SUMMARY:", json.dumps(total["summary"], ensure_ascii=False))


if __name__ == "__main__":
    main()
