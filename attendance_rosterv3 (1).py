from datetime import datetime
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from mfrc522 import SimpleMFRC522
import RPi.GPIO as GPIO
from RPLCD.i2c import CharLCD

SHEET_NAME = "Attendance Log"
TAB_NAME = "Attendance Log"
CREDS_FILE = "credentials.json"

# Set this to your class tardy cutoff time using 24-hour format.
# Example: "09:30" = 9:30 AM, "13:05" = 1:05 PM
TARDY_TIME = "09:30"

SCAN_COOLDOWN = 5
BUZZER_PIN = 21

last_uid = None
last_scan_time = 0

GPIO.setwarnings(False)
GPIO.setmode(GPIO.BCM)
GPIO.setup(BUZZER_PIN, GPIO.OUT)
GPIO.output(BUZZER_PIN, GPIO.LOW)

# LCD setup
lcd = CharLCD(
    i2c_expander='PCF8574',
    address=0x27,
    port=1,
    cols=16,
    rows=2
)

def lcd_write(line1="", line2=""):
    lcd.clear()
    lcd.write_string(line1[:16])
    if line2:
        lcd.cursor_pos = (1, 0)
        lcd.write_string(line2[:16])

def buzz(duration=0.15):
    GPIO.output(BUZZER_PIN, GPIO.HIGH)
    time.sleep(duration)
    GPIO.output(BUZZER_PIN, GPIO.LOW)

def buzz_read_ack():
    # Immediate feedback that the RFID card was detected/read.
    buzz(0.04)

def buzz_success():
    buzz(0.08)
    time.sleep(0.05)
    buzz(0.08)

def buzz_unknown():
    buzz(0.4)

def buzz_already_scanned():
    buzz(0.08)
    time.sleep(0.05)
    buzz(0.08)
    time.sleep(0.05)
    buzz(0.08)

def connect_to_sheet():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, scope)
    client = gspread.authorize(creds)
    sheet = client.open(SHEET_NAME).worksheet(TAB_NAME)
    return sheet

def setup_headers(sheet):
    headers = ["Name", "ID #", "Status", "Time In"]
    current = sheet.row_values(1)
    if current != headers:
        sheet.update("A1:D1", [headers])

def get_status():
    now_time = datetime.now().time()
    tardy_cutoff = datetime.strptime(TARDY_TIME, "%H:%M").time()

    if now_time > tardy_cutoff:
        return "Tardy"
    else:
        return "Present"

def shorten_name(full_name):
    parts = full_name.split()
    if len(parts) >= 2:
        return f"{parts[0]} {parts[-1][0]}."
    return full_name

def find_student_row_by_uid(sheet, uid):
    uid_text = str(uid)
    try:
        cell = sheet.find(uid_text)
        return cell.row
    except gspread.exceptions.CellNotFound:
        return None

reader = SimpleMFRC522()

try:
    print("Connecting to Google Sheet...")
    lcd_write("Connecting...", "")
    sheet = connect_to_sheet()
    setup_headers(sheet)
    print("Ready. Scan a card.")
    lcd_write("Ready", "Scan Card")

    while True:
        uid, text = reader.read()
        now = time.time()

        if uid == last_uid and (now - last_scan_time) < SCAN_COOLDOWN:
            print(f"Ignored duplicate scan: {uid}")
            time.sleep(1)
            continue

        last_uid = uid
        last_scan_time = now

        # Immediate feedback as soon as the card is read.
        buzz_read_ack()

        print("Scanned UID:", uid)

        row = find_student_row_by_uid(sheet, uid)

        if row is None:
            print("Unknown Card")
            lcd_write("Unknown Card", str(uid))
            buzz_unknown()
            time.sleep(2)
            lcd_write("Ready", "Scan Card")
            print("-----")
            continue

        name = sheet.cell(row, 1).value
        existing_status = (sheet.cell(row, 3).value or "").strip()
        existing_time = sheet.cell(row, 4).value or ""
        lcd_name = shorten_name(name)

        # Reset behavior:
        # If Status is Absent, the card is allowed to scan again.
        # If Status is Present or Tardy, the scan is blocked as already checked in.
        if existing_status in ["Present", "Tardy"]:
            print("Name:", name)
            print("Already scanned.")
            print("Existing Status:", existing_status)
            print("Existing Time In:", existing_time)

            lcd_write(lcd_name, "Already Scanned")
            buzz_already_scanned()

            time.sleep(2)
            lcd_write("Ready", "Scan Card")

            print("Sheet not updated.")
            print("-----")
            time.sleep(1)
            continue

        status = get_status()
        log_time = datetime.now().strftime("%I:%M %p").lstrip("0")

        print("Name:", name)
        print("Status:", status)
        print("Time In:", log_time)

        # Update existing roster row:
        # Column C = Status
        # Column D = Time In
        sheet.update(f"C{row}:D{row}", [[status, log_time]])

        lcd_write(lcd_name, f"{status} {log_time}")

        if status == "Tardy":
            buzz_unknown()
        else:
            buzz_success()

        time.sleep(2)
        lcd_write("Ready", "Scan Card")

        print("Updated sheet.")
        print("-----")
        time.sleep(1)

except KeyboardInterrupt:
    print("Program stopped.")
    lcd_write("Program", "Stopped")

finally:
    lcd.clear()
    GPIO.cleanup()
