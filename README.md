import cv2
from pyzbar.pyzbar import decode
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import time
import json

# Google Sheets setup
scope = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file("qr.json")
client = gspread.authorize(creds.with_scopes(scope))

# Open your Google Sheet
sheet_id = "1GSOkG9dGjH76OdMTDJsym8QvVTDuOaXJhhhWHf-yI6c"
spreadsheet = client.open_by_key(sheet_id)
worksheet = spreadsheet.sheet1  # Opens the first sheet/tab in the spreadsheet

# Start camera
cap = cv2.VideoCapture(0)

# Set a delay time between scans (in seconds)
scan_delay = 2
last_scanned_data = None

while True:
    _, frame = cap.read()
    qr_codes = decode(frame)

    if qr_codes:
        for barcode in qr_codes:
            qr_data = barcode.data.decode('utf-8')

            # Check if the data is different from the last scanned data
            if qr_data != last_scanned_data:
                now = datetime.now()
                date = now.strftime('%Y-%m-%d')
                time_str = now.strftime('%H:%M:%S')

                try:
                    # Parse the JSON data
                    qr_data_dict = json.loads(qr_data)

                    # Extract the values from the dictionary
                    name = qr_data_dict.get("name", "")
                    roll = qr_data_dict.get("roll", "")
                    position = qr_data_dict.get("position", "")

                    # Append data to Google Sheet
                    worksheet.append_row([date, time_str, name, roll, position])

                     # Update last scanned data
                    last_scanned_data = qr_data

                    print(f"Scanned QR Code: {qr_data} on {date} at {time_str}")

                    # Show the scanned QR code data on the frame
                    cv2.putText(frame, f"Scanned: {qr_data}", (barcode.rect.left, barcode.rect.top - 10),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 0, 0), 2)

                    # Introduce a delay after a successful scan
                    time.sleep(scan_delay)

                except json.JSONDecodeError:
                    print("Failed to decode QR data as JSON. Skipping.")

    # Display the frame
    cv2.imshow('QR Code Scanner', frame)

    # Press 'q' to quit
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
