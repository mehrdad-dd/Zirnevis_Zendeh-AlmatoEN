import time
import re
import os

import ctranslate2
import sentencepiece as spm
from pywinauto import Desktop
from pywinauto.findwindows import ElementNotFoundError


# -----------------------------
# CONFIG
# -----------------------------
MODEL_DIR = r"D:/Projects/sentence_transformers/translation_models/JustFrederik_nllb-200-distilled-600M-ct2-int8"
SRC_LANG = "deu_Latn"
TGT_LANG = "eng_Latn"

POLL_WINDOW_SEC = 0.35
POLL_CAPTION_SEC = 0.20

STABILITY_SECONDS = 0.7
MIN_LEN_NO_PUNCT = 18

# Excel files (DIRECT PATHS ON DRIVE D)
EXCEL_FILES = [
    r"D:\Personal\MDD\fani.xlsx",
    r"D:\Personal\MDD\hr.xlsx",
]

# -----------------------------
# ANSI COLOR HELPER
# -----------------------------
def color(text, code):
    return f"\x1b[{code}m{text}\x1b[0m"


# -----------------------------
# UI TEXT FILTERS
# -----------------------------
UI_BLACKLIST_EXACT = {
    "Position",
    "Sprache ändern",
    "Einstellungen",
    "Weitere Informationen",
    "Filtern von Obszönitäten",
    "Mikrofonaudio einschließen",
    "Untertitelstil",
}

UI_BLACKLIST_CONTAINS = [
    "untertitel", "titel", "live",
    "einstellungen", "informationen",
    "mikrofon", "obszön",
    "sprache", "position",
    "untertitelstil",
]


# -----------------------------
# MODEL INIT
# -----------------------------
translator = ctranslate2.Translator(
    MODEL_DIR,
    device="cpu",
    compute_type="int8",
    inter_threads=1,
    intra_threads=0
)

sp = spm.SentencePieceProcessor()
sp.load(f"{MODEL_DIR}/sentencepiece.bpe.model")


# -----------------------------
# STATE (dedupe + stability)
# -----------------------------
last_printed_norm = ""
pending_text = ""
pending_norm = ""
pending_since = 0.0

excel_opened_once = False


# -----------------------------
# TEXT HELPERS
# -----------------------------
def clean_line(text: str):
    text = (text or "").strip()
    if not text:
        return None

    if text in UI_BLACKLIST_EXACT:
        return None

    low = text.lower()
    if any(b in low for b in UI_BLACKLIST_CONTAINS):
        return None

    if len(text) < 3:
        return None

    text = text.split("\n")[-1].strip()
    text = re.sub(r"(.)\1{3,}", r"\1", text)
    return text


def norm_for_dedupe(s: str) -> str:
    s = (s or "").strip().lower()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[.?!…]+$", "", s).strip()
    return s


# -----------------------------
# LIVE WINDOW HELPERS
# -----------------------------
def find_live_window(desktop: Desktop):
    try:
        win = desktop.window(title_re=r".*Liveuntertitel.*")
        if win.exists(timeout=0.15):
            return win
    except Exception:
        pass
    return None


def is_any_menu_open(window) -> bool:
    try:
        return len(window.descendants(control_type="MenuItem")) > 0
    except Exception:
        return False


def get_latest_caption_line(window):
    if is_any_menu_open(window):
        return None

    try:
        texts = window.descendants(control_type="Text")
    except Exception:
        return None

    candidates = []
    for t in texts:
        try:
            s = (t.window_text() or "").strip()
            if not s:
                continue

            if s in UI_BLACKLIST_EXACT:
                continue

            low = s.lower()
            if any(b in low for b in UI_BLACKLIST_CONTAINS):
                continue

            last = s.split("\n")[-1].strip()
            if len(last) < 3:
                continue

            r = t.rectangle()
            score = (r.bottom * 1000) + min(len(last), 250)
            candidates.append((score, last))
        except Exception:
            continue

    if not candidates:
        return None

    candidates.sort(key=lambda x: x[0])
    return candidates[-1][1]


# -----------------------------
# EXCEL OPEN
# -----------------------------
def open_excel_files_once(filepaths):
    # Try COM first
    try:
        import win32com.client  # type: ignore
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        for p in filepaths:
            if os.path.exists(p):
                excel.Workbooks.Open(p)
                print(color("✔ Opened Excel:", "32"), p)
            else:
                print(color("⚠ File not found:", "31"), p)
        return
    except Exception:
        pass

    # Fallback
    for p in filepaths:
        if os.path.exists(p):
            os.startfile(p)
            print(color("✔ Opened:", "32"), p)
        else:
            print(color("⚠ File not found:", "31"), p)


# -----------------------------
# MAIN LOOP
# -----------------------------
desktop = Desktop(backend="uia")
print("Waiting for Liveuntertitel...")
print("Listening (DE -> EN)...")

while True:
    try:
        window = find_live_window(desktop)
        if not window:
            time.sleep(POLL_WINDOW_SEC)
            continue

        # Open Excel files ONCE after services are up
        if not excel_opened_once:
            open_excel_files_once(EXCEL_FILES)
            excel_opened_once = True

        latest = get_latest_caption_line(window)
        if not latest:
            time.sleep(POLL_CAPTION_SEC)
            continue

        cleaned = clean_line(latest)
        if not cleaned:
            time.sleep(POLL_CAPTION_SEC)
            continue

        now = time.time()
        n = norm_for_dedupe(cleaned)

        if not n:
            time.sleep(POLL_CAPTION_SEC)
            continue

        if n == last_printed_norm:
            time.sleep(POLL_CAPTION_SEC)
            continue

        if n != pending_norm:
            pending_norm = n
            pending_text = cleaned
            pending_since = now
            time.sleep(POLL_CAPTION_SEC)
            continue

        if (now - pending_since) < STABILITY_SECONDS:
            time.sleep(POLL_CAPTION_SEC)
            continue

        stable_text = pending_text

        if len(stable_text) < MIN_LEN_NO_PUNCT and not re.search(r"[.!?]$", stable_text):
            time.sleep(POLL_CAPTION_SEC)
            continue

        last_printed_norm = pending_norm

        source_tokens = [SRC_LANG] + sp.encode(stable_text, out_type=str) + ["</s>"]

        res = translator.translate_batch(
            [source_tokens],
            target_prefix=[[TGT_LANG]],
            end_token="</s>",
            beam_size=2,
            max_decoding_length=80,
            repetition_penalty=1.25,
            no_repeat_ngram_size=3
        )

        out = res[0].hypotheses[0]
        if out and out[0] == TGT_LANG:
            out = out[1:]
        if out and out[-1] == "</s>":
            out = out[:-1]

        translated = sp.decode(out)

        print("DE:", stable_text)
        print(color("EN:", "36"), color(translated, "36"))
        print("-" * 60)

    except ElementNotFoundError:
        time.sleep(0.3)
    except Exception as e:
        print(color("Error:", "31"), e)

    time.sleep(POLL_CAPTION_SEC)
