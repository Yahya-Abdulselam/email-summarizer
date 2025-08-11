# app.py
# Streamlit Email Summarizer with OAuth (Gmail + Microsoft 365) and optional Outlook Redemption
# -----------------------------------------------------------
# Requirements (see notes below):
# streamlit, requests, msal, google-auth, google-auth-oauthlib, google-api-python-client,
# beautifulsoup4, transformers, torch (optional), pywin32 (optional), bs4

import os, json, base64, re, datetime as dt, uuid, requests
import streamlit as st

# --- COM imports (Windows only, optional Redemption fallback) ---
try:
    import win32com.client as win32
    import pywintypes
except Exception:
    win32 = None
    pywintypes = None

from bs4 import BeautifulSoup

# ================== UI / CONFIG ==================
st.set_page_config(page_title="Daily Email Summarizer", layout="centered")
st.title("📬  Daily Email Summarizer")

# Read secrets / config
APP_BASE_URL = st.secrets.get("APP_BASE_URL", "")  # must match your Streamlit URL
GOOGLE_CLIENT_ID = st.secrets.get("GOOGLE_CLIENT_ID", "")
GOOGLE_CLIENT_SECRET = st.secrets.get("GOOGLE_CLIENT_SECRET", "")
MS_CLIENT_ID = st.secrets.get("MS_CLIENT_ID", "")
MS_CLIENT_SECRET = st.secrets.get("MS_CLIENT_SECRET", "")
MS_TENANT = st.secrets.get("MS_TENANT", "common")  # "common" or your tenant ID
# Scopes
GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.readonly", "openid", "email", "profile"]
MS_SCOPES = ["Mail.Read", "offline_access", "openid", "profile"]

# Basic controls
N = st.slider("Number of latest emails", 1, 50, 10)

DIGEST_LINK_MODE = st.selectbox(
    "Links in Daily Digest",
    ["describe", "shorten", "remove", "keep"],
    index=0,
    help="How links appear in the Daily Digest list."
)

DETAILS_LINK_MODE = st.selectbox(
    "Links in Details (summaries)",
    ["phrase-link", "shorten-link", "keep", "remove"],
    index=0,
    help="phrase-link = small description as a clickable link; shorten-link = domain as link."
)

# OAuth session state
if "google_token" not in st.session_state:
    st.session_state.google_token = None
if "ms_token" not in st.session_state:
    st.session_state.ms_token = None
if "oauth_nonce" not in st.session_state:
    st.session_state.oauth_nonce = str(uuid.uuid4())

# ================== Helpers (text) ==================
def strip_html(html: str) -> str:
    if not html:
        return ""
    return BeautifulSoup(html, "html.parser").get_text(separator=" ", strip=True)

QUOTE_MARKERS = [
    "-----Original Message-----","From:","Sent:","To:","Cc:","Subject:",
    "On ","wrote:","发件人","De :","Van:","Von:"
]

def dequote_and_designature(text: str) -> str:
    text = (text or "").strip()
    if not text:
        return ""
    lines, out = text.splitlines(), []
    for ln in lines:
        if any(m in ln for m in QUOTE_MARKERS): break
        if ln.strip().startswith(("--","—")): break
        if ln.strip().lower().startswith(("best,","regards,","thanks,")): break
        out.append(ln)
    t = "\n".join(out)
    t = re.sub(r'\n{2,}','\n\n',t)
    t = re.sub(r'(This email.*confidential.*)','',t, flags=re.I|re.S)
    return t.strip()

def first_sentences(text: str, max_sent=5):
    sents = [s.strip() for s in re.split(r'(?<=[.!?])\s+', text) if s.strip()]
    sents = [s for s in sents if len(s) > 30][:max_sent]
    return sents or ([text[:120].strip()] if text else [])

# ---------- Link normalization ----------
_URL_RE = re.compile(
    r'(?P<pre>[<\(\[])?(?P<url>(?:https?://|www\.)[^\s<>)\]]+)(?P<post>[>\)\]])?',
    re.I
)

def _domain_hint(dom: str):
    d = dom.lower()
    if "zoom.us" in d: return "Zoom meeting link"
    if "teams.microsoft" in d or "teams.live" in d: return "Microsoft Teams link"
    if "meet.google" in d: return "Google Meet link"
    if "calendar.google" in d or ("outlook.office.com" in d and "calendar" in d): return "calendar event"
    if any(x in d for x in ["drive.google", "dropbox.com", "box.com", "sharepoint.com", "onedrive"]): return "cloud file"
    if "docs.google" in d: return "Google Doc"
    if "sheets.google" in d: return "Google Sheet"
    if "slides.google" in d: return "Google Slides"
    if "github.com" in d: return "GitHub link"
    if "atlassian.net" in d or "jira" in d: return "Jira ticket"
    if "figma.com" in d: return "Figma file"
    if "notion.so" in d: return "Notion page"
    return None

def normalize_links(text: str, mode: str = "describe", keep_url: bool = False) -> str:
    if not text:
        return text
    if (not keep_url and mode == "remove") or (keep_url and mode == "remove"):
        text = re.sub(r'\[([^\]]+)\]\(\s*https?://[^\s)]+\s*\)', r'\1', text, flags=re.I)
    def repl(m):
        url = m.group("url")
        dom = re.sub(r'^https?://', '', url, flags=re.I).split('/')[0]
        if keep_url:
            if mode == "keep":
                return url
            if mode == "remove":
                return ""
            if mode == "shorten-link":
                return f"[{dom}]({url})"
            label = _domain_hint(dom) or dom
            return f"[{label}]({url})"
        else:
            if mode == "keep":
                return url
            if mode == "remove":
                return ""
            if mode == "shorten":
                return dom
            return (_domain_hint(dom) or dom)
    out = _URL_RE.sub(repl, text)
    out = re.sub(r'\s*([<\(\[])\s*([>\)\]])\s*', ' ', out)
    out = re.sub(r'\s{2,}', ' ', out).strip()
    return out

def extract_urls(text: str):
    if not text:
        return []
    urls, seen = [], set()
    for m in _URL_RE.finditer(text):
        u = m.group("url")
        if u not in seen:
            urls.append(u); seen.add(u)
    return urls

@st.cache_resource(show_spinner=False)
def get_summarizer():
    try:
        from transformers import pipeline
        try:
            import torch
            device = 0 if torch.cuda.is_available() else -1
        except Exception:
            device = -1
        return pipeline("summarization", model="sshleifer/distilbart-cnn-12-6", device=device)
    except Exception:
        return None

def summarize_text(text: str):
    text = (text or "").strip()
    if not text:
        return []
    text = text[:6000]
    summ = get_summarizer()
    if summ:
        try:
            out = summ(text, max_length=130, min_length=60, do_sample=False)[0]["summary_text"]
            return first_sentences(out, max_sent=5)[:5]
        except Exception:
            pass
    return first_sentences(text, max_sent=5)

def safe_str(x):
    try:
        return str(x)
    except Exception:
        return ""

# ================== OAuth utilities ==================
def _b64url(d: dict) -> str:
    raw = json.dumps(d, separators=(",", ":")).encode("utf-8")
    return base64.urlsafe_b64encode(raw).decode("ascii")

def _unb64url(s: str) -> dict:
    try:
        data = base64.urlsafe_b64decode(s + "==")
        return json.loads(data)
    except Exception:
        return {}

def _current_url_without_params():
    return APP_BASE_URL  # we assume Streamlit served at this base

# ---------- Google OAuth ----------
def google_auth_url():
    auth_endpoint = "https://accounts.google.com/o/oauth2/v2/auth"
    redirect_uri = _current_url_without_params()
    state = _b64url({"p": "google", "nonce": st.session_state.oauth_nonce})
    params = {
        "client_id": GOOGLE_CLIENT_ID,
        "response_type": "code",
        "redirect_uri": redirect_uri,
        "scope": " ".join(GMAIL_SCOPES),
        "access_type": "offline",
        "prompt": "consent",
        "include_granted_scopes": "true",
        "state": state,
    }
    from urllib.parse import urlencode
    return f"{auth_endpoint}?{urlencode(params)}"

def google_fetch_token(code: str):
    token_endpoint = "https://oauth2.googleapis.com/token"
    redirect_uri = _current_url_without_params()
    data = {
        "code": code,
        "client_id": GOOGLE_CLIENT_ID,
        "client_secret": GOOGLE_CLIENT_SECRET,
        "redirect_uri": redirect_uri,
        "grant_type": "authorization_code",
    }
    r = requests.post(token_endpoint, data=data, timeout=30)
    r.raise_for_status()
    return r.json()  # includes access_token, refresh_token (if any), expires_in, id_token

def google_refresh_token(refresh_token: str):
    token_endpoint = "https://oauth2.googleapis.com/token"
    data = {
        "client_id": GOOGLE_CLIENT_ID,
        "client_secret": GOOGLE_CLIENT_SECRET,
        "refresh_token": refresh_token,
        "grant_type": "refresh_token",
    }
    r = requests.post(token_endpoint, data=data, timeout=30)
    r.raise_for_status()
    return r.json()

# ---------- Microsoft OAuth (MSAL) ----------
def ms_auth_url():
    from msal import ConfidentialClientApplication
    authority = f"https://login.microsoftonline.com/{MS_TENANT}"
    redirect_uri = _current_url_without_params()
    cca = ConfidentialClientApplication(
        client_id=MS_CLIENT_ID,
        authority=authority,
        client_credential=MS_CLIENT_SECRET,
    )
    state = _b64url({"p": "ms", "nonce": st.session_state.oauth_nonce})
    return cca.get_authorization_request_url(
        scopes=MS_SCOPES,
        state=state,
        redirect_uri=redirect_uri,
        prompt="select_account",
        response_mode="query",
    )

def ms_fetch_token(code: str):
    from msal import ConfidentialClientApplication
    authority = f"https://login.microsoftonline.com/{MS_TENANT}"
    redirect_uri = _current_url_without_params()
    cca = ConfidentialClientApplication(
        client_id=MS_CLIENT_ID,
        authority=authority,
        client_credential=MS_CLIENT_SECRET,
    )
    result = cca.acquire_token_by_authorization_code(
        code=code,
        scopes=MS_SCOPES,
        redirect_uri=redirect_uri,
    )
    if "access_token" not in result:
        raise RuntimeError(f"MS token error: {result}")
    return result

def ms_refresh_token(refresh_token: str):
    from msal import ConfidentialClientApplication
    authority = f"https://login.microsoftonline.com/{MS_TENANT}"
    cca = ConfidentialClientApplication(
        client_id=MS_CLIENT_ID,
        authority=authority,
        client_credential=MS_CLIENT_SECRET,
    )
    result = cca.acquire_token_by_refresh_token(refresh_token, scopes=MS_SCOPES)
    if "access_token" not in result:
        raise RuntimeError(f"MS refresh error: {result}")
    return result

# ================== Email fetchers ==================
def gmail_fetch_latest(n=10, token=None):
    """Return list of dicts: from, subject, received, body, body_original, links"""
    if not token:
        return []
    access = token.get("access_token")
    if not access:
        return []

    headers = {"Authorization": f"Bearer {access}"}
    # List latest messages
    r = requests.get(
        "https://gmail.googleapis.com/gmail/v1/users/me/messages",
        params={"maxResults": n, "q": "", "labelIds": "INBOX"},
        headers=headers, timeout=30
    )
    r.raise_for_status()
    ids = [m["id"] for m in r.json().get("messages", [])]

    out = []
    for mid in ids:
        r = requests.get(
            f"https://gmail.googleapis.com/gmail/v1/users/me/messages/{mid}",
            params={"format": "full"},
            headers=headers, timeout=30
        )
        r.raise_for_status()
        data = r.json()

        headers_list = data.get("payload", {}).get("headers", [])
        hdr = {h["name"].lower(): h["value"] for h in headers_list}
        frm = hdr.get("from", "")
        subj = hdr.get("subject", "")
        dtm = hdr.get("date", "")

        # best-effort body extraction (prefers text/plain, fallback to text/html)
        def _walk_parts(p):
            if not p: return []
            if p.get("mimeType", "").startswith("multipart/"):
                parts = p.get("parts", []) or []
                out = []
                for child in parts:
                    out.extend(_walk_parts(child))
                return out
            else:
                return [p]

        parts = _walk_parts(data.get("payload", {}))
        body_text, body_html = "", ""
        for p in parts:
            mt = p.get("mimeType", "")
            b64 = (p.get("body", {}) or {}).get("data", "")
            if not b64:
                continue
            try:
                raw = base64.urlsafe_b64decode(b64 + "==").decode("utf-8", errors="ignore")
            except Exception:
                raw = ""
            if mt == "text/plain" and not body_text:
                body_text = raw
            elif mt == "text/html" and not body_html:
                body_html = raw
        raw_body = body_text or strip_html(body_html)

        body = dequote_and_designature(raw_body)
        if not body:
            continue

        out.append({
            "from": frm,
            "subject": subj,
            "received": dtm,
            "body": body,
            "body_original": raw_body,
            "links": extract_urls(raw_body),
        })
    return out

def ms_fetch_latest(n=10, token=None):
    """Return list of dicts from Microsoft Graph."""
    if not token:
        return []
    access = token.get("access_token")
    if not access:
        return []

    headers = {"Authorization": f"Bearer {access}"}
    # Select a few fields; include body and bodyPreview; order by receivedDateTime desc
    url = "https://graph.microsoft.com/v1.0/me/messages"
    params = {
        "$top": str(n),
        "$orderby": "receivedDateTime desc",
        "$select": "sender,from,receivedDateTime,subject,body,bodyPreview"
    }
    r = requests.get(url, headers=headers, params=params, timeout=30)
    r.raise_for_status()
    items = r.json().get("value", [])

    out = []
    for m in items:
        frm = (m.get("from") or {}).get("emailAddress", {}).get("name") or \
              (m.get("from") or {}).get("emailAddress", {}).get("address") or ""
        subj = m.get("subject", "") or ""
        dtm = m.get("receivedDateTime", "") or ""
        body_content = ((m.get("body") or {}).get("content") or "")
        body_raw = strip_html(body_content) if (m.get("body") or {}).get("contentType") == "html" else (body_content or "")
        raw_body = body_raw or (m.get("bodyPreview") or "")

        body = dequote_and_designature(raw_body)
        if not body:
            continue

        out.append({
            "from": frm,
            "subject": subj,
            "received": dtm,
            "body": body,
            "body_original": raw_body,
            "links": extract_urls(raw_body),
        })
    return out

# --------- Redemption (RDO) optional ----------
def redemption_available() -> bool:
    if not win32:
        return False
    try:
        _ = win32.Dispatch("Redemption.RDOSession")
        return True
    except Exception:
        return False

def connect_outlook_rdo(profile_name: str = None):
    if not win32:
        st.error("Windows COM not available. Run on Windows with pywin32 installed.")
        return None, None
    if not redemption_available():
        st.error(
            "Redemption is not installed/registered on this machine. "
            "Install the Redemption MSI or register Redemption.dll."
        )
        return None, None
    try:
        rdo = win32.Dispatch("Redemption.RDOSession")
        if profile_name:
            rdo.Logon(profile_name)
        else:
            rdo.Logon()

        inbox = rdo.GetDefaultFolder(6)  # olFolderInbox
        items = inbox.Items
        try:
            items.MAPITable.Sort(0x0E060040, True)  # PR_MESSAGE_DELIVERY_TIME desc
        except Exception:
            try:
                items.Sort("[ReceivedTime]", True)
            except Exception:
                pass
        return items, None
    except Exception as e:
        return None, f"Redemption RDO connection failed: {e}"

def redemption_fetch_latest(n=10, items=None):
    if not items:
        return []
    emails, count = [], min(n, items.Count)
    i, fetched = 1, 0
    while fetched < count and i <= items.Count:
        try:
            itm = items.Item(i); i += 1
            if itm is None:
                continue
            body_raw = safe_str(getattr(itm, "Body", "")) \
                       or strip_html(safe_str(getattr(itm, "HTMLBody", "")))
            body = dequote_and_designature(body_raw)
            if not body:
                continue
            emails.append({
                "from": safe_str(getattr(itm, "SenderName","") or getattr(itm, "SenderEmailAddress","")),
                "subject": safe_str(getattr(itm, "Subject","")),
                "received": safe_str(getattr(itm, "ReceivedTime", dt.datetime.now())),
                "body": body,
                "body_original": body_raw,
                "links": extract_urls(body_raw)
            })
            fetched += 1
        except Exception:
            continue
    return emails

# ================== OAuth callback handling ==================
# Read query params for OAuth redirects (both providers return to the same APP_BASE_URL)
query_params = st.query_params  # Streamlit 1.30+; for older versions use experimental_get_query_params()
code = query_params.get("code", [None])[0] if isinstance(query_params.get("code"), list) else query_params.get("code")
state = query_params.get("state", [None])[0] if isinstance(query_params.get("state"), list) else query_params.get("state")
error = query_params.get("error", [None])[0] if isinstance(query_params.get("error"), list) else query_params.get("error")

if error:
    st.error(f"OAuth error: {error}")

if code and state:
    decoded = _unb64url(state)
    prov = decoded.get("p")
    nonce = decoded.get("nonce")
    if nonce != st.session_state.oauth_nonce:
        st.warning("State mismatch. Ignoring OAuth response.")
    else:
        try:
            if prov == "google":
                tok = google_fetch_token(code)
                st.session_state.google_token = tok
                st.success("Connected to Google (Gmail).")
            elif prov == "ms":
                tok = ms_fetch_token(code)
                st.session_state.ms_token = tok
                st.success("Connected to Microsoft 365 (Outlook).")
            # clear code/state from URL
            st.query_params.clear()
        except Exception as e:
            st.error(f"Token exchange failed: {e}")

# ================== Sign-in UI ==================
st.subheader("Accounts")

cols = st.columns(3)
with cols[0]:
    if st.session_state.google_token:
        if st.button("🔌 Disconnect Google"):
            st.session_state.google_token = None
        else:
            st.caption("✅ Google connected")
    else:
        if GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET:
            st.link_button("🔐 Sign in with Google", google_auth_url())
        else:
            st.caption("Add GOOGLE_* secrets to enable Gmail")

with cols[1]:
    if st.session_state.ms_token:
        if st.button("🔌 Disconnect Microsoft"):
            st.session_state.ms_token = None
        else:
            st.caption("✅ Microsoft connected")
    else:
        if MS_CLIENT_ID and MS_CLIENT_SECRET:
            st.link_button("🔐 Sign in with Microsoft", ms_auth_url())
        else:
            st.caption("Add MS_* secrets to enable Outlook via Graph")

with cols[2]:
    if win32:
        st.caption("Optional: Outlook (local profile via Redemption)")
    else:
        st.caption("Windows COM not available")

# ================== Fetch + Summarize ==================
go = st.button("Summarize")

if go:
    all_emails = []

    # Gmail
    if st.session_state.google_token:
        tok = st.session_state.google_token
        # refresh if needed (very simple heuristic)
        if "refresh_token" in tok:
            # try, but don't crash if refresh fails
            try:
                refreshed = google_refresh_token(tok["refresh_token"])
                tok.update({k: v for k, v in refreshed.items() if k != "refresh_token"})
                st.session_state.google_token = tok
            except Exception:
                pass
        try:
            emails = gmail_fetch_latest(n=N, token=st.session_state.google_token)
            for e in emails:
                e["_source"] = "Gmail"
            all_emails.extend(emails)
        except Exception as e:
            st.warning(f"Gmail fetch failed: {e}")

    # Microsoft Graph
    if st.session_state.ms_token:
        tok = st.session_state.ms_token
        if "refresh_token" in tok:
            try:
                refreshed = ms_refresh_token(tok["refresh_token"])
                # MS returns a whole new token dict; keep refresh_token
                if "refresh_token" not in refreshed and "refresh_token" in tok:
                    refreshed["refresh_token"] = tok["refresh_token"]
                st.session_state.ms_token = refreshed
            except Exception:
                pass
        try:
            emails = ms_fetch_latest(n=N, token=st.session_state.ms_token)
            for e in emails:
                e["_source"] = "Outlook (Graph)"
            all_emails.extend(emails)
        except Exception as e:
            st.warning(f"Microsoft fetch failed: {e}")

    # Redemption (optional)
    if win32 and st.toggle("Also include local Outlook (Redemption)", value=False):
        items, err = connect_outlook_rdo(profile_name=None)
        if err:
            st.warning(err)
        elif items:
            emails = redemption_fetch_latest(n=N, items=items)
            for e in emails:
                e["_source"] = "Outlook (local)"
            all_emails.extend(emails)

    if not all_emails:
        st.info("No recent emails found or no accounts connected.")
    else:
        # Sort newest-first if possible by received timestamp
        def _parse_dt(v):
            try:
                return dt.datetime.fromisoformat(v.replace("Z","+00:00"))
            except Exception:
                try:
                    # Gmail Date header fallback
                    from email.utils import parsedate_to_datetime
                    return parsedate_to_datetime(v)
                except Exception:
                    return dt.datetime.min
        all_emails.sort(key=lambda e: _parse_dt(e.get("received","")), reverse=True)

        st.subheader("🗞️ Daily Digest")
        for idx, e in enumerate(all_emails, 1):
            bullets = summarize_text(e["body"])
            bullets = [normalize_links(b, DIGEST_LINK_MODE, keep_url=False) for b in bullets]
            one_liner = bullets[0] if bullets else (e["subject"] or "No subject")
            src = e.get("_source","")
            st.markdown(f"{idx}. **{e['subject'] or 'No subject'}** — {one_liner}  <small>· {src}</small>", unsafe_allow_html=True)

        st.markdown("---")
        st.subheader("🔎 Details")
        for idx, e in enumerate(all_emails, 1):
            bullets = summarize_text(e["body"])
            if DETAILS_LINK_MODE in ("phrase-link", "shorten-link"):
                bullets = [normalize_links(b, DETAILS_LINK_MODE, keep_url=True) for b in bullets]
            elif DETAILS_LINK_MODE == "remove":
                bullets = [normalize_links(b, "remove", keep_url=False) for b in bullets]
            with st.expander(f"{idx}. {e['subject'] or 'No subject'}  ·  {e['from']}  ·  {e['received']}  ·  {e.get('_source','')}"):
                st.markdown("**Summary:**")
                for b in bullets[:5]:
                    st.markdown(f"- {b}")

                if e.get("links"):
                    if DETAILS_LINK_MODE in ("phrase-link", "shorten-link"):
                        links_line = ", ".join(normalize_links(u, DETAILS_LINK_MODE, keep_url=True) for u in e["links"])
                    elif DETAILS_LINK_MODE == "keep":
                        links_line = ", ".join(e["links"])
                    else:
                        links_line = ""
                    if links_line:
                        st.markdown("**Links:** " + links_line)

                st.markdown("**Body (preview):**")
                preview = normalize_links(e["body_original"], "shorten", keep_url=False)
                st.code((preview[:500] + ("…" if len(preview) > 500 else "")))
else:
    st.caption("All processing stays in memory. OAuth tokens stay in session. "
               "You can connect Gmail and/or Outlook (Graph). Optional Windows Redemption fallback.")
