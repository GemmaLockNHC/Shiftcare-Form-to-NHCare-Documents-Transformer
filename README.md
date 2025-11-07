# JotForm Service Agreement Generator

A web application that generates Service Agreement PDFs and CSV exports from uploaded PDF forms.

## Quick Start

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Run the web app:
   ```bash
   python3 app.py
   ```

3. Open your browser to `http://localhost:5000`

## Required Files

All required input files are in `outputs/other/`:
- `NDIS Support Items - NDIS Support Items.csv` - Support item pricing data
- `Active_Users_1761707021.csv` - Staff contact information

## Features

- Upload PDF forms and extract data
- Generate Service Agreement PDFs
- Generate Client Export CSV files
- Choose to generate one or both outputs

## Deployment

See `DEPLOYMENT_CHECKLIST.md` for deployment instructions.
