# *** This file text has been sanitized by AI for spell checking and to prevent gramatic mistakes. ***

# Liveuntertitel-DE-to-EN
Live German-to-English Subtitle Translator

🎧 Live German-to-English Subtitle Translator

A real-time subtitle extraction and translation tool that captures German live captions (Liveuntertitel) from Windows UI and translates them into English using an optimized NLLB-200 (CTranslate2) model.

🚀 Features
🔴 Real-time subtitle monitoring from Windows Live Caption window
🇩🇪 ➡️ 🇺🇸 German → English translation
⚡ Fast inference using ctranslate2 with INT8 optimization
🧠 Smart filtering & deduplication to avoid noisy or repeated outputs
🧹 UI noise removal (menus, settings, etc.)
📊 Auto-opens Excel files for workflow integration
🎯 Stable caption detection (avoids incomplete sentences)
🧠 How It Works
Connects to Windows UI using pywinauto
Detects the Liveuntertitel window
Extracts the latest caption text
Cleans and filters UI noise
Waits for text stability
Translates using:
ctranslate2
sentencepiece
Prints both:
Original German 🇩🇪
Translated English 🇺🇸
🏗️ Requirements

Install dependencies:

pip install ctranslate2 sentencepiece pywinauto pywin32
📦 Model Setup

You need a converted CTranslate2 model (INT8 recommended):

Example:

JustFrederik/nllb-200-distilled-600M (converted to CT2)

Update path in config:

MODEL_DIR = "D:/path/to/your/model"
⚙️ Configuration

Main settings:

SRC_LANG = "deu_Latn"
TGT_LANG = "eng_Latn"

POLL_WINDOW_SEC = 0.35
POLL_CAPTION_SEC = 0.20

STABILITY_SECONDS = 0.7
MIN_LEN_NO_PUNCT = 18
📊 Excel Integration

Automatically opens Excel files at startup:

EXCEL_FILES = [
    r"D:\Personal\MDD\fani.xlsx",
    r"D:\Personal\MDD\hr.xlsx",
]
▶️ Usage
Enable Windows Live Captions (German)
Run the script:
python main.py
Output example:
DE: Ich habe heute viel gearbeitet
EN: I worked a lot today
------------------------------------------------------------
🧹 Filtering Logic

The script intelligently removes:

UI elements like:
Settings
Language menu
Microphone options
Short/noisy text
Repeated captions
Unstable partial sentences
⚡ Performance Notes
Uses INT8 quantization → faster CPU inference
Low latency polling (~200ms)
Beam search optimized for speed/quality balance
⚠️ Limitations
Works only on Windows
Requires Liveuntertitel window
UI changes in Windows may break detection
Translation quality depends on model
🛠️ Possible Improvements
GUI overlay for subtitles
Multi-language support
GPU acceleration
Export to file / database
Streaming API support
📄 License

MIT License (or your preferred license)

❤️ Credits
Meta AI — NLLB-200
CTranslate2
SentencePiece
PyWinAuto
