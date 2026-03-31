# pdf-review-workflow
"you can batch messy real-world PDFs and still trust the output because of review"

> ⚠️ This is NOT a PDF parser.
>
> This project is a **framework for working with PDF parsing in an operations perspective**.

---

https://github.com/user-attachments/assets/d0085fcf-6ed8-49a4-9688-df7066e84a11

## What this is

This project provides a simple way to:

- take PDFs directly from email (including Outlook drag & drop)
- process multiple files at once
- handle multiple suppliers in the same batch
- generate a **review PDF with overlays** for validation
- convert everything into a single Excel file

It is designed to run as a **local tool inside a company**, not as a SaaS or generic converter.

---

## What this is NOT

This project does **NOT solve PDF parsing**.

PDF parsing is:
- supplier-specific
- format-dependent
- inherently unreliable

👉 You are expected to:
- test different parsing approaches
- implement your own parsers
- choose what works best for your specific documents

The included parsers (e.g. SodaAntarctica / BigCustomer) are **examples only**.

---

## Intended use case

This tool is useful when:

- you receive structured documents via email (e.g. order confirmations, invoices)
- you do NOT have EDI
- you get documents in batches ("in bursts")
- you want to quickly move data into Excel for further processing

---

## Typical flow

1. Drag emails directly from Outlook into the web UI  
2. System extracts PDF attachments  
3. Each PDF is parsed using a supplier-specific parser  
4. A review PDF is generated (with overlays)  
5. Operator verifies extracted data  
6. Data is used in Excel  

---

## Key idea

Separate the problem into two parts:

### 1. Parsing (hard problem)

You implement:
- supplier detection
- field extraction
- data interpretation

### 2. Workflow (this project)

This project handles:
- file intake (PDF / .msg / .eml)
- batch processing  
- review visualization (overlay on original PDF)
- Excel aggregation  
- job queue + worker system  

---

## Why this approach works

Instead of trying to build a “perfect parser”, this system:

- accepts that parsing will fail sometimes
- makes errors visible via review PDFs
- lets operators verify results quickly
- allows mixing multiple suppliers in one run

---


## Review step (important)

Every processed PDF is included in a **review document**:

- overlays show extracted values directly on the original PDF
- errors and unknown suppliers are clearly marked
- nothing is silently dropped

This is critical because:

> PDF parsing is not deterministic — verification is required.

---
<img width="961" height="451" alt="image" src="https://github.com/user-attachments/assets/c0528d4c-7f42-4b8e-b444-0ac35f8158e1" />

Example of the review step:

- Yellow fields = extracted data  
- Red boxes = source locations in the PDF  
- "Excelrow" corresponds to the row in the output Excel file  

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

## Limitations

- Only single-page PDFs are supported (by design)
- Parsing is fully custom per supplier
- No guarantee of correctness without review
- Not designed for cloud or scaling (local use)

---

## Final note

This project is about:

making PDF parsing _usable_, not perfect, in practice.

