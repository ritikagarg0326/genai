# ✅ How to Install Dependencies

Follow these simple steps to get the Reddit Scraper running!

---

## Step 1 — Make sure Python is installed

Open a terminal (Command Prompt on Windows) and type:

```
python --version
```

You should see something like `Python 3.9.0` or higher.
If not, download Python from: https://www.python.org/downloads/

---

## Step 2 — Install all dependencies

In the same terminal, navigate to this folder and run:

```
pip install -r requirements.txt
```

That's it! This will download and install:
- `requests` — used to fetch data from Reddit
- `openpyxl` — used to create the Excel (.xlsx) report file

---

## Step 3 — Run the app

```
python main.py
```

A colourful window will pop up — fill in the form and click **GENERATE REPORT**! 🚀

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `pip` not found | Try `python -m pip install -r requirements.txt` |
| `python` not found | Try `python3` instead of `python` |
| Permission error | Add `--user` flag: `pip install -r requirements.txt --user` |

---

> 💡 **Note:** `tkinter` (the GUI library) comes bundled with Python — no need to install it separately!
