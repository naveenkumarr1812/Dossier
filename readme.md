# Dossier — Employee Document Organizer

A Streamlit app that reads a folder of unorganized employee documents (Aadhar,
PAN, Bank Details, Payslips, Employment Letters — mixed PDFs/photos, random
filenames, even rotated/photocopied scans), identifies who each document
belongs to using a Groq **vision** model, and lets you build one clean folder
per employee.

## Setup

```bash
pip install -r requirements.txt
```

You need **Poppler** installed so PDF pages can be rendered to images before
being sent to the vision model (only needed if you have PDF files — plain
JPG/PNG photos work without it):
https://github.com/oschwartz10612/poppler-windows/releases
Extract it and add the `bin` folder to your PATH, or set `POPPLER_PATH` at
the top of `app.py`.

You'll also need a **Groq API key** — https://console.groq.com/keys
(paste it into the app's password field, or set env var `GROQ_API_KEY`
before running).

> **Important — models get deprecated.** `GROQ_VISION_MODEL` at the top of
> `app.py` points to a vision-capable Groq model. Groq regularly retires
> models with only weeks of notice (e.g. `llama-3.3-70b-versatile` was
> retired in June 2026) — if every file starts coming back "UNCLEAR"/"Unknown"
> after previously working, that's usually a silently-failing API call from
> a deprecated model name, not a real extraction problem. Check
> console.groq.com/docs/models for the current vision-capable model and
> update the constant. The app now surfaces the actual error text next to
> each file in the extraction log, so this is easier to spot going forward.

## Run

```bash
streamlit run app.py
```

## How it works

**Step 01 — Extract & Identify**
Upload every file from your messy documents folder in one go. Each file's
image (or first PDF page, rendered to an image) is sent directly to a Groq
vision model, which reads the employee name, document type, and ID number —
even from rotated, upside-down, or photocopied scans — and returns
structured JSON. Since the same person's name can appear slightly
differently across documents, a fuzzy name-matching pass clusters variants
together (adjustable via the "Name match strictness" slider). You get an
editable table, a live per-file log (including any API errors), and a
downloadable Excel file to review by hand.

**Step 02 — Verify & Organize**
Open the Excel, fix anything wrong (misspelled names, wrong document column,
UNMATCHED rows), save it, and upload it back. The app matches the filenames
in each row back to the files it stored during Step 01, copies them into a
folder per employee, and gives you a ZIP to download.

## Notes

- Files are stored server-side under `./workspace/<session_id>/` for the
  duration of your session — delete this folder whenever you like once
  you've downloaded the ZIP.
- If you come back to Step 02 in a new session (server restarted, etc.), the
  app will ask you to re-upload the original documents — it matches purely
  by filename at that point, so keep filenames unchanged in the Excel.
- If one photographed page contains multiple documents (e.g. Aadhar front
  and back pasted on the same sheet), the model is prompted to identify the
  primary/most complete document on that page — it won't split one image
  into two rows.
- For large batches (100+ files), extraction will take a while since each
  file is a separate LLM call — this is a good spot to add batching/caching
  if you need it later.