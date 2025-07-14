import pandas as pd
from fpdf import FPDF
import os
import smtplib
import mimetypes
from email.message import EmailMessage
import platform
import subprocess
from datetime import datetime
from pathlib import Path
import pdfkit

# Mapping sheet names to hall display names
hall_mapping = {
    "GR": "G.R FUNCTION HALL",
    "MINI": "G.R MINI FUNCTION HALL",
    "Gardens": "G.R GARDENS"
}

# Prompt for sheet selection
print("Available halls: GR, MINI, Gardens")
selected_sheet = input("Enter which hall (GR / MINI / Gardens): ").strip()

if selected_sheet not in hall_mapping:
    print("Invalid sheet name.")
    exit()

# Load Excel file with selected sheet
excel_path = "bookings (2).xlsx"
df = pd.read_excel(excel_path, sheet_name=selected_sheet)

# Prompt for Booking ID
booking_id = input("Enter Booking ID to generate bill: ")
booking = df[df["BookingID"] == int(booking_id)]

if booking.empty:
    print("No booking found with that ID.")
    exit()
else:
    row = booking.iloc[0]

    # Extract and format values
    hall_name = hall_mapping[selected_sheet]
    client_name = row["ClientName"]
    address = row["Address"]
    phone = row["PhoneNo"]
    email = row["email"]
    event = row["Event"]
    id_proof = row["IDProof"]
    booking_date = pd.to_datetime(row["BookingDate"]).strftime("%d-%m-%Y")
    event_date = pd.to_datetime(row["BookedDate"]).strftime("%d-%m-%Y")

    rent = int(str(row["Rent"]).replace(',', ''))
    units = int(row["UnitsConsumed"])
    rate = int(row["RatePerUnit"])
    electricity = units * rate
    total = rent + electricity

    def number_to_words(n):
        import inflect
        p = inflect.engine()
        return p.number_to_words(n, andword="")

    total_in_words = number_to_words(total).upper()

    # Generate HTML content
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; }}
            .bill {{ border: 4px solid black; padding: 20px; width: 95%; margin: 0 auto; }}
            .top-section {{ display: flex; justify-content: space-between; margin-bottom: 20px; }}
            .centered-header {{ text-align: center; font-size: 20px; font-weight: bold; position: relative; }}
            .bold {{ font-weight: bold; }}
            .amount-box {{ border: 2px solid #000; padding: 10px; width: 100%; margin-top: 20px; }}
            .amount-box td {{ padding: 5px; }}
            .amount-box .label {{ text-align: left; }}
            .amount-box .value {{ text-align: right; }}
        </style>
    </head>
    <body>
        <div class="bill">
            <div class="centered-header">
                CASH / VOUCHER<br>
                {hall_name}<br>
                # 1-850, Gongadi Ramappa Compound, 3rd Road, ANANTHAPURAMU<br><br>
                <span style="position: absolute; top: 0; right: 0; font-size: 14px;">STD : 08554-273141</span>
            </div>
            <br><hr><br>
            <div class="top-section">
                <div>
                    <div><span class="bold">Booking ID:</span> {booking_id}</div>
                    <div><span class="bold">Phone:</span> {phone}</div>
                </div>
                <div><span class="bold">Date:</span> {booking_date}</div>
            </div>

            <div style="margin-bottom: 10px;"><em><span class="bold">Received with thanks from</span> {client_name}</em></div>
            <div style="margin-bottom: 10px;"><em>{address}</em></div>
            <div style="margin-bottom: 20px;"><em>Towards {event} on {event_date}</em></div>

            <table class="amount-box">
                <tr>
                    <td class="label bold">Rent</td>
                    <td class="value"><b>Rs. {rent}</b></td>
                </tr>
                <tr>
                    <td><hr></td>
                </tr>
                <tr>
                    <td class="label bold">Electricity</td>
                    <td></td>
                </tr>
                <tr>
                    <td class="label">Consumed Units {units} x {rate}</td>
                    <td class="value"><b>Rs. {electricity}</b></td>
                </tr>
                <tr>
                    <td class="label bold">Total</td>
                    <td class="value bold"><b>Rs.</b> {total}</td>
                </tr>
            </table>

            <div style="margin-top: 20px;">
                (<span class="bold">{total_in_words} RUPEES ONLY</span>)
            </div>
        </div>
    </body>
    </html>
    """

    # Save HTML to file
    html_filename = f"Bill_{booking_id}.html"
    with open(html_filename, "w", encoding="utf-8") as f:
        f.write(html_content)

    # Convert HTML to PDF with wkhtmltopdf config
    pdf_filename = f"Bill_{booking_id}.pdf"
    config = pdfkit.configuration(wkhtmltopdf=r"C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe")
    try:
        if os.path.exists(pdf_filename):
            os.remove(pdf_filename)
        pdfkit.from_file(html_filename, pdf_filename, configuration=config)
    except Exception as e:
        print("PDF generation failed:", e)
        exit()

    # Open PDF
    if platform.system() == "Windows":
        os.startfile(pdf_filename)
    elif platform.system() == "Darwin":
        subprocess.run(["open", pdf_filename])
    else:
        subprocess.run(["xdg-open", pdf_filename])

    # Email (if email provided)
    if pd.notna(email) and email.strip() != "":
        sender_email = "gongadiakhilesh@gmail.com" # Replace with your email
        # Use an app password for Gmail if 2FA is enabled
        app_password = "xxxxxxxxxxxxxxxxxxxx" # Replace with your app password
        receiver_email = email
        cc_email = "gongadiakhilesh@gmail.com"# Replace with CC email

        msg = EmailMessage()
        msg["Subject"] = f"Bill for Booking ID {booking_id}"
        msg["From"] = sender_email
        msg["To"] = receiver_email
        msg["Cc"] = cc_email
        msg.set_content("Please find attached your bill for the booking.")

        with open(pdf_filename, "rb") as f:
            file_data = f.read()
            maintype, subtype = mimetypes.guess_type(pdf_filename)[0].split('/')
            msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=pdf_filename)

        try:
            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.starttls()
                server.login(sender_email, app_password)
                server.send_message(msg)
                print(f"Email sent to {receiver_email} and CC to {cc_email}")
        except Exception as e:
            print("Failed to send email:", e)
    else:
        # Print if no email
        if platform.system() == "Windows":
            os.startfile(pdf_filename, "print")
        elif platform.system() == "Darwin":
            subprocess.run(["lp", pdf_filename])
        else:
            subprocess.run(["lpr", pdf_filename])
