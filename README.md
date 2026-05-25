# Chapter 10 Results PDF Generator

Splits Google Form responses into 3 clean Satsang Vihar result PDFs:
**Balika**, **ND1**, **ND2**.

---

## 📦 What's in this folder

```
Chapter10_PDF_Generator/
├── split_responses.py     ← the script that does the work
├── responses.xlsx         ← your form data (REPLACE this every week)
├── requirements.txt       ← list of Python libraries to install
├── fonts/
│   └── NotoSansGujarati.ttf  ← Gujarati font (already included)
└── README.md              ← this file
```

---

## 🚀 ONE-TIME SETUP (5 minutes)

### Step 1 — Install Python
Download from https://www.python.org/downloads/

⚠️ **CRITICAL:** On the first install screen, tick the checkbox
**"Add Python to PATH"** at the bottom. If you skip this, nothing will work.

### Step 2 — Install VS Code
Download from https://code.visualstudio.com/

### Step 3 — Open this folder in VS Code
- Open VS Code
- File → **Open Folder** → select this `Chapter10_PDF_Generator` folder
- If asked "Do you trust the authors?" → click Yes

### Step 4 — Install the Python extension in VS Code
- Click the **Extensions icon** on the left (4 squares) or press `Ctrl+Shift+X`
- Search for **"Python"**
- Install the one by **Microsoft** (top result, blue checkmark)

### Step 5 — Install the libraries
- Open VS Code's terminal: press `` Ctrl+` `` (Ctrl + backtick)
- In the terminal, type this and press Enter:

```
pip install -r requirements.txt
```

- Wait ~30 seconds. You should see "Successfully installed ..." at the end.

✅ **Setup done. You'll never need to do this again.**

---

## 📅 EVERY WEEK (30 seconds)

### Step 1 — Replace responses.xlsx
1. Download the new responses from Google Forms (File → Download → .xlsx)
2. Rename it to **`responses.xlsx`** (exactly this name)
3. Drag it into this folder, replacing the old one

### Step 2 — Run the script
In VS Code:
- Open `split_responses.py` (click it in the left sidebar)
- Click the **▶ Play button** in the top-right corner of the editor

OR in the terminal:
```
python split_responses.py
```

### Step 3 — Get your PDFs
A new **`output/`** folder appears with 3 PDFs:
- `Satsang_Vihar_Paper_3_Balika.pdf`
- `Satsang_Vihar_Paper_3_ND1.pdf`
- `Satsang_Vihar_Paper_3_ND2.pdf`

Right-click any PDF → **"Reveal in File Explorer"** to find it on your disk.

---

## ❓ Troubleshooting

**`python` is not recognized**
→ Python wasn't added to PATH. Reinstall Python and tick the PATH checkbox.

**`pip install` shows errors**
→ Try `python -m pip install -r requirements.txt` instead.

**`ERROR: 'responses.xlsx' not found`**
→ The file must be named exactly `responses.xlsx` (lowercase, no spaces) and
   placed in the same folder as `split_responses.py`.

**Columns changed in my Google Form**
→ Open `split_responses.py` and edit the `COL_*` variables near the top.
   They're clearly labeled.

**Score isn't out of 20**
→ Edit `MAX_SCORE = 20` near the top of the script.

---

## 📝 What the script does

For each row in `responses.xlsx`, it figures out the group based on which
name field is filled in:

| If this column has a value | Goes to PDF |
|---|---|
| `બાલિકાનું નામ` (column E) | Balika |
| `બાળકનું નામ` (column H) | ND1 |
| `બાળકનું નામ 2` (column K) | ND2 |

Each PDF has the same layout as the official Satsang Vihar Online Test
Result : Paper 3 reports.

That's it. Enjoy. 🙏
