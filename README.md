# TyranoScript Localization Kit

A toolkit for translating TyranoScript visual novel games.  
Extracts text from `.ks` files into an Excel spreadsheet, translates it via machine translation, and writes the translated text back into new `.ks` files.

---

## Quick Start

1. **Install Python 3.8+** from [python.org](https://www.python.org/downloads/) (make sure to check "Add to PATH" during installation)
2. **Download** and unzip this kit anywhere on your computer
3. **Double-click** `launcher.bat` to open the GUI
4. **Set the scenario path** — click "Browse..." and navigate to your game's `data/scenario` folder
5. **Click the three buttons in order:**
   - **1. Extract** — pulls all translatable text from `.ks` files into `translations.xlsx`
   - **2. Translate** — machine-translates the spreadsheet (progress is saved; you can pause/resume or close and continue later)
   - **3. Apply** — writes translations back into new `.ks` files in the `translated/` folder
6. **Copy** the translated `.ks` files back into your game's scenario folder (make a backup first!)

---

## Requirements

- **Python 3.8+** (tested on 3.10 - 3.14)
- **pip packages:** `openpyxl`, `deep-translator` (installed automatically on first run)

To install manually:
```
pip install openpyxl deep-translator
```

---

## Translation Engines

| Engine     | Quality           | Limit                | API Key  |
|------------|-------------------|----------------------|----------|
| **google** | Good (default)    | Unlimited            | Not needed |
| **deepl**  | Best for European languages | 500K chars/month | Free key from [deepl.com/pro-api](https://www.deepl.com/pro-api) |
| **libre**  | Decent            | Unlimited            | Not needed |
| **mymemory** | Basic           | 5,000 chars/day      | Not needed |

**Recommendation:** For Russian, German, French, and other European languages, DeepL gives the best results. Register for a free API key at [deepl.com/pro-api](https://www.deepl.com/pro-api), then paste it into the "DeepL API key" field in the launcher.

---

## Files in This Kit

| File                | Description                                                   |
|---------------------|---------------------------------------------------------------|
| `launcher.bat`      | Double-click to start the GUI                                 |
| `launcher.pyw`      | GUI application (tkinter)                                     |
| `config.bat`        | Project settings for .bat files (auto-updated by the GUI)     |
| `config.json`       | Settings file (auto-generated, stores your last configuration)|
| `tyrano_l10n.py`    | Core script: extract `.ks` -> XLSX and apply XLSX -> `.ks`    |
| `translate_xlsx.py` | Machine translation of XLSX via Google / DeepL / Libre / MyMemory |
| `01_extract.bat`    | Command-line alternative: extract strings                     |
| `02_translate.bat`  | Command-line alternative: translate (prompts for engine/language) |
| `03_apply.bat`      | Command-line alternative: apply translations                  |

---

## Detailed Guide

### Step 1: Extract

Scans all `.ks` files in the scenario folder and creates `translations.xlsx` with the following columns:

| Column | Content                                      |
|--------|----------------------------------------------|
| A      | File name                                     |
| B      | Line number                                   |
| C      | Type (text / name / choice / etc.)            |
| D      | Tag (original TyranoScript tag, if any)        |
| E      | **Original** — the source text                |
| F      | **Translation** — fill this in (or let Step 2 do it) |

You can edit `translations.xlsx` in Excel or LibreOffice before or after machine translation to fix errors manually.

### Step 2: Translate

Reads column E (Original) and fills column F (Translation) using the selected engine.

- **Progress is saved** every 50 rows to `translations_translated.xlsx`
- **Pause/Resume**: click the Pause button in the GUI to pause translation between batches
- **Interrupt safe**: if you close the window, progress is saved. Next time you run translation, it picks up where it left off (existing translations are not overwritten)
- **Change engine mid-way**: you can start with Google, pause, then switch to DeepL and resume

### Step 3: Apply

Reads `translations_translated.xlsx` and creates new `.ks` files in the `translated/` folder with the original text replaced by translations.

**Important:** Always back up your original scenario files before copying the translated ones over them!

---

## Command-Line Usage

You can also run the scripts directly without the GUI:

```bash
# Extract
python tyrano_l10n.py extract "C:\path\to\scenario" --output translations.xlsx

# Translate (Google, default)
python translate_xlsx.py translations.xlsx --from en --to ru

# Translate (DeepL)
python translate_xlsx.py translations.xlsx --engine deepl --deepl-key YOUR_KEY --from en --to ru

# Translate (LibreTranslate)
python translate_xlsx.py translations.xlsx --engine libre --from en --to de

# List available languages
python translate_xlsx.py --list --engine google

# Apply translations
python tyrano_l10n.py apply "C:\path\to\scenario" translations_translated.xlsx --output translated
```

---

## FAQ

**Q: Translation is slow. Can I speed it up?**  
A: Reduce the delay between batches: `--delay 0.2` (command line) or edit `DELAY_BETWEEN` in `translate_xlsx.py`. Be careful — too fast may trigger rate limiting.

**Q: Some strings shouldn't be translated (names, tags, etc.)**  
A: Open `translations.xlsx` in Excel, put the original text into the Translation column for those rows (so they stay unchanged), then run Translate — it will skip rows that already have translations.

**Q: I want to translate into multiple languages**  
A: Run Extract once, then make copies of the XLSX for each language. Run Translate separately for each copy with different `--to` values.

**Q: The translated `.ks` files have encoding issues**  
A: The tool preserves the original file encoding. If you see issues, make sure your original files are UTF-8.

**Q: Can I edit translations after machine translation?**  
A: Yes! Open `translations_translated.xlsx` in Excel, edit column F as needed, save, then run Apply.

---

## License

MIT License. Free to use, modify, and distribute.
