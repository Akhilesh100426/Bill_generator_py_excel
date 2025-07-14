# Bill_generator_py_excel
# 🧾 Function Hall Booking Bill Generator

This project automates the generation of a professional PDF bill (cash/voucher) for bookings made at different function halls. The bill includes details like client information, rent, electricity charges, total cost (in numbers and words), and can be emailed to the client directly.

---

## 📂 Files Included

- **bill.py**: Python script to generate a PDF bill from booking data.
- **bookings (2).xlsx**: Excel file containing booking details for 3 halls — GR, MINI, and Gardens.
- **Bill_101.pdf**: Sample output PDF bill generated for Booking ID 101.

---

## 🚀 Features

- Generate bills from Excel data based on booking ID.
- Converts booking details into a formatted HTML invoice.
- Converts the HTML bill to PDF using `wkhtmltopdf`.
- Automatically opens or prints the PDF.
- Emails the PDF to the client if an email is provided.
- Outputs the total amount in both digits and words.

---

## 🛠️ Prerequisites

- Python 3.7+
- Required libraries:
  - `pandas`
  - `fpdf`
  - `pdfkit`
  - `inflect`
- External:
  - `wkhtmltopdf` (Install from [https://wkhtmltopdf.org/](https://wkhtmltopdf.org/))
  - SMTP credentials for sending email (Gmail supported)

Install Python dependencies using:
```bash
pip install pandas fpdf inflect pdfkit
