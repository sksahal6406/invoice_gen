# Diamond Pharma Tax Invoice — Streamlit

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

The app generates:
- PDF in the same landscape tax-invoice layout shown in the supplied reference image.
- Excel workbook with the same invoice structure and A4 landscape print settings.

The seller/company information and fixed invoice text are hard-coded from the supplied invoice. Customer, invoice, product and adjustment fields are editable.

Important: because the reference provided is a photograph of a printed invoice rather than the original blank PDF/Excel template, the generated document is a vector recreation of the visible format. If you provide the original blank invoice template, the coordinates can be matched much more precisely.
