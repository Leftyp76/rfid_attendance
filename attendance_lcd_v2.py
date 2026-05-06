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

STUDENTS = {
    435804206226: "Hayden P. Perhach",
    91334669530: "Jonah V. Small",
    503106401882: "Misty M. Perhach"
}

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

def buzz_success():
    buzz(0.08)
    time.sleep(0.05)
    buzz(0.08)

def buzz_unknown():
    buzz(0.4)

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
    headers = ["Timestamp", "ID #", "Name", "Status"]
    current = sheet.row_values(1)
    if current != headers:
        sheet.update("A1:D1", [headers])

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

        timestamp = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")

        if uid in STUDENTS:
            name = STUDENTS[uid]
        else:
            name = "Unknown Card"

        print("Scanned UID:", uid)
        print("Name:", name)
        lcd_write("Scanned:", name)

        sheet.append_row([timestamp, str(uid), name])
        print("Logged to sheet.")

        # Show logged time on LCD
        log_time = datetime.now().strftime("%I:%M %p").lstrip("0")
        lcd_write(name, f"Logged {log_time}")

        if name == "Unknown Card":
            print("Buzzing unknown")
            buzz_unknown()
        else:
            print("Buzzing success")
            buzz_success()

        time.sleep(2)
        lcd_write("Ready", "Scan Card")

        print("-----")
        time.sleep(1)

except KeyboardInterrupt:
    print("Program stopped.")
    lcd_write("Program", "Stopped")

finally:
    lcd.clear()
    GPIO.cleanup()
