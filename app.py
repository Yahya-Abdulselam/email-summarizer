# app.py
# Streamlit Outlook Daily Email Summarizer (Windows + Outlook Desktop) via Redemption RDO
# Usage:
#   1) Install Redemption (MSI or DLL registration)
#   2) pip install streamlit transformers torch sentencepiece beautifulsoup4 pywin32
#   3) streamlit run app.py

import re, datetime as dt
import streamlit as st

# --- COM imports (Windows only) ---
try:
    import win32com.client as win32
    import pywintypes
except Exception:
    win32 = None
    pywintypes = None

from bs4 import BeautifulSoup

# ---------------- UI ----------------
st.set_page_config(page_title="Daily Email Summarizer", layout="centered")
st.title("📬 Easiest Daily Email Summarizer")

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

go = st.button("Summarize")

# ------------- Helpers -------------
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
# eats optional wrappers (<...>, (...), [...] ) so no stray "<>"
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
    """
    If keep_url=False (Daily Digest): modes = describe | shorten | remove | keep  (plain text)
    If keep_url=True  (Details):      modes = phrase-link | shorten-link | keep | remove (Markdown links)
    - Consumes wrappers (<...>, (...), [...]) around bare URLs.
    - Preserves Markdown anchor text unless removing.
    """
    if not text:
        return text

    # If removing, also drop Markdown-style links entirely ([text](url) -> text)
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
            # phrase-link (default)
            label = _domain_hint(dom) or dom
            return f"[{label}]({url})"
        else:
            if mode == "keep":
                return url
            if mode == "remove":
                return ""
            if mode == "shorten":
                return dom
            # describe (default)
            return (_domain_hint(dom) or dom)

    out = _URL_RE.sub(repl, text)
    # Remove any empty wrapper pairs that might still linger
    out = re.sub(r'\s*([<\(\[])\s*([>\)\]])\s*', ' ', out)
    out = re.sub(r'\s{2,}', ' ', out).strip()
    return out

def extract_urls(text: str):
    """Return unique URLs in order of appearance."""
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

# --------- Redemption (RDO) ----------
def redemption_available() -> bool:
    if not win32:
        return False
    try:
        _ = win32.Dispatch("Redemption.RDOSession")
        return True
    except Exception:
        return False

def connect_outlook_rdo(profile_name: str = None):
    """
    Connect using Redemption RDO (Extended MAPI, no Outlook security prompts).
    If multiple Outlook profiles exist and you want a specific one, pass profile_name.
    """
    if not win32:
        st.error("Windows COM not available. Run on Windows with pywin32 installed.")
        return None, None
    if not redemption_available():
        st.error(
            "Redemption is not installed/registered on this machine. "
            "Install the Redemption MSI or register Redemption.dll (see notes below)."
        )
        return None, None
    try:
        rdo = win32.Dispatch("Redemption.RDOSession")
        if profile_name:
            rdo.Logon(profile_name)
        else:
            rdo.Logon()  # default profile

        inbox = rdo.GetDefaultFolder(6)  # 6 == olFolderInbox
        items = inbox.Items  # RDOItems

        # Sort newest first using MAPITable (fast + reliable)
        try:
            # PR_MESSAGE_DELIVERY_TIME = 0x0E060040, True => descending
            items.MAPITable.Sort(0x0E060040, True)
        except Exception:
            try:
                items.Sort("[ReceivedTime]", True)
            except Exception:
                pass

        return items, None
    except Exception as e:
        return None, f"Redemption RDO connection failed: {e}"

def safe_str(x):
    try:
        return str(x)
    except Exception:
        return ""

# --------------- Main ----------------
if go:
    PROFILE_NAME = None  # e.g., "Outlook" or your org-specific name

    items, err = connect_outlook_rdo(profile_name=PROFILE_NAME)
    if err:
        st.warning(err)
    elif items is None:
        pass
    else:
        emails, count = [], min(N, items.Count)
        i, fetched = 1, 0

        # Iterate newest-first
        while fetched < count and i <= items.Count:
            try:
                itm = items.Item(i); i += 1  # RDOMail (1-based index)
                if itm is None:
                    continue

                # Prefer plain text body; fall back to HTML stripped
                body_raw = safe_str(getattr(itm, "Body", "")) \
                           or strip_html(safe_str(getattr(itm, "HTMLBody", "")))

                # Clean quotes/signatures but KEEP URLs so Details can link them
                body = dequote_and_designature(body_raw)
                if not body:
                    continue

                emails.append({
                    "from": safe_str(getattr(itm, "SenderName","") or getattr(itm, "SenderEmailAddress","")),
                    "subject": safe_str(getattr(itm, "Subject","")),
                    "received": safe_str(getattr(itm, "ReceivedTime", dt.datetime.now())),
                    "body": body,                 # cleaned for summarization
                    "body_original": body_raw,    # original with links
                    "links": extract_urls(body_raw)
                })
                fetched += 1

            except Exception:
                continue

        if not emails:
            st.info("No recent emails found to summarize.")
        else:
            st.subheader("🗞️ Daily Digest")
            for idx, e in enumerate(emails, 1):
                bullets = summarize_text(e["body"])
                # Daily Digest: text-only link handling
                bullets = [normalize_links(b, DIGEST_LINK_MODE, keep_url=False) for b in bullets]
                one_liner = bullets[0] if bullets else (e["subject"] or "No subject")
                st.markdown(f"{idx}. **{e['subject'] or 'No subject'}** — {one_liner}")

            st.markdown("---")
            st.subheader("🔎 Details")
            for idx, e in enumerate(emails, 1):
                bullets = summarize_text(e["body"])
                # Details: clickable links if chosen
                if DETAILS_LINK_MODE in ("phrase-link", "shorten-link"):
                    bullets = [normalize_links(b, DETAILS_LINK_MODE, keep_url=True) for b in bullets]
                elif DETAILS_LINK_MODE == "remove":
                    bullets = [normalize_links(b, "remove", keep_url=False) for b in bullets]
                # else "keep": leave bullets as-is

                with st.expander(f"{idx}. {e['subject'] or 'No subject'}  ·  {e['from']}  ·  {e['received']}"):
                    st.markdown("**Summary:**")
                    for b in bullets[:5]:
                        st.markdown(f"- {b}")

                    # If the model didn't copy URLs into the bullets, still show them
                    if e.get("links"):
                        if DETAILS_LINK_MODE in ("phrase-link", "shorten-link"):
                            links_line = ", ".join(normalize_links(u, DETAILS_LINK_MODE, keep_url=True) for u in e["links"])
                        elif DETAILS_LINK_MODE == "keep":
                            links_line = ", ".join(e["links"])
                        else:  # remove
                            links_line = ""
                        if links_line:
                            st.markdown("**Links:** " + links_line)

                    st.markdown("**Body (preview):**")
                    # readable preview without raw URLs littering it
                    preview = normalize_links(e["body_original"], "shorten", keep_url=False)
                    st.code((preview[:500] + ("…" if len(preview) > 500 else "")))
else:
    st.caption("All processing stays in memory. No passwords, IMAP, or disk writes. Uses your logged-in Outlook profile via Extended MAPI (Redemption).")
