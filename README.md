# pdf-review-engine

Batch messy real-world PDFs — and still trust the output.

A review-first engine for PDF parsing.

---
> ⚠️ This is NOT a PDF parser.
>
> This project **makes PDF parsing usable** in real operations.

---

https://github.com/user-attachments/assets/d0085fcf-6ed8-49a4-9688-df7066e84a11

## Typical flow (video demo)

1. Drag emails directly from Outlook into the web UI  
2. System extracts PDF attachments  
3. Each PDF is parsed using a supplier-specific parser  
4. A review PDF is generated (with overlays)  
5. Operator verifies extracted data  
6. Data is moved to Excel  


It is designed to run as a **local tool inside a company**, not as a SaaS or generic converter.

---

## What this is NOT

This project does **NOT solve PDF parsing**.

PDF parsing is:
- supplier-specific
- format-dependent
- unreliable

👉 You are expected to:
- test different parsing approaches
- implement your own parsers
- choose what works best for your specific documents

The included parsers (e.g. SodaAntarctica / BigCustomer) are **examples only**.

---

## Intended use case

- You receive structured PDFs via email (orders, invoices, etc.)
- You do not have EDI
- Documents arrive in batches
- You need to move data into Excel — reliably

---

## Key idea

Separate the problem into two parts:

### 1. Parsing (hard problem)

You implement:
- supplier detection
- field extraction
- data interpretation

### 2. Engine (this project)

This project handles:
- file intake (PDF / .msg / .eml)
- batch processing  
- review visualization (overlay on original PDF)
- Excel aggregation  
- job queue + worker system  

---
<img width="961" height="451" alt="image" src="https://github.com/user-attachments/assets/c0528d4c-7f42-4b8e-b444-0ac35f8158e1" />

Example of the review step:

- Yellow fields = extracted data  
- Red boxes = source locations in the PDF  
- "Excelrow" links each extracted value to its row in the output Excel file

---

## The (very important) review step

This system is built around one assumption:

> parsing will fail.

Every processed PDF is turned into a **review document**:

- extracted values are overlaid on the original PDF
- errors are visible immediately
- nothing is silently dropped

This enables:

- fast human verification
- safe batch processing across suppliers

Instead of chasing perfect parsing,
this system makes imperfect parsing usable.

---


## Architecture

### Web app
- Upload UI (Flask + Dropzone)  
- Job creation and status polling  

### Worker
- Picks up queued jobs  
- Processes files  
- Produces:
  - Review PDF  
  - Excel output  

---

## Setup

Install dependencies:

```bash
pip install -r requirements.txt
```



Run web app and the worker:
```
python flask_app.py
python worker.py
```

---

## Compatibility

Tested with:
- Chrome, Edge, Firefox  
- Windows and Linux  
- Outlook drag & drop (.msg files)

---

## Positioning

Most PDF tools focus on extracting data.

This project is built around reviewing that data.

Every result is turned into a visual review document,
where extracted values are overlaid on the original PDF
and explicitly linked to the output.

Instead of trusting parsing,
you verify it.

---

## Limitations

- Only single-page PDFs are supported (by design)
- Parsing is fully custom per supplier
- No guarantee of correctness without review
- Not designed for cloud or scaling (local use)

---

## Philosophy

Make PDF parsing usable — not perfect.

