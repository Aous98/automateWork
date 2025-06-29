import openpyxl
import pyautogui
import time
from pynput import keyboard
from datetime import datetime
import pyperclip

# Global variables
should_continue = True
patient_codes = []

def wait_for_backspace():
    """Wait for Backspace key press to continue"""
    def on_press(key):
        if key == keyboard.Key.backspace:
            return False
    with keyboard.Listener(on_press=on_press) as listener:
        listener.join()

def capture_patient_code():
    """Copy current cell content (patient code)"""
    for _ in range(2):  # Try twice to ensure copy works
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.3)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.3)
        try:
            code = pyperclip.paste().strip()
            if code:
                patient_codes.append(code)
                print(f"Captured patient code: {code}")
                return
        except:
            pass
    print("Could not capture patient code")
    patient_codes.append("Unknown")

def format_excel_value(value):
    """Convert Excel values to proper strings"""
    if isinstance(value, datetime):
        return value.strftime("%m/%d/%Y")  # Format dates as MM/DD/YYYY
    elif value is None:
        return ""
    else:
        # Handle numbers, strings, and special characters
        return str(value).strip()

def type_safely(text):
    """Reliable typing with error handling"""
    try:
        # First try normal typing
        pyautogui.write(text, interval=0.1)  # Slower typing for reliability
    except:
        try:
            # Fallback to clipboard paste
            pyperclip.copy(text)
            time.sleep(0.3)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.3)
        except:
            print(f"Failed to type: {text[:20]}...")

def process_patients(file_path):
    global should_continue, patient_codes
    
    # Load Excel workbook
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    
    print("Position cursor in FIRST NAME field and press BACKSPACE to start...")
    wait_for_backspace()
    time.sleep(1)  # Initial delay
    
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not should_continue:
            break
            
        # Process all fields with error handling
        fields = [
            (row[0], 2),   # First Name + 2 tabs
            (row[1], 2),   # Last Name + 2 tabs
            (None, 3),      # Patient Code (handled separately)
            (row[2], 2),    # Address + 2 tabs
            (row[3], 3),    # City + 3 tabs
            (row[4], 11),   # Zip + 11 tabs
            (row[5], 1),    # Gender + 1 tab
            (row[6], 0),    # Birthday (no tab after)
        ]
        
        # Type personal info
        for value, tabs in fields:
            if value is not None:  # Skip patient code position
                type_safely(format_excel_value(value))
            pyautogui.press('tab', presses=tabs, interval=0.1)
        
        # Capture patient code
        capture_patient_code()
        
        # Insurance info
        print("Press BACKSPACE to continue to insurance info...")
        wait_for_backspace()
        
        insurance_fields = [
            (row[7], 4),   # Payer + 4 tabs
            (row[8], 0),    # Member ID (no tab after)
        ]
        
        for value, tabs in insurance_fields:
            type_safely(format_excel_value(value))
            pyautogui.press('tab', presses=tabs, interval=0.1)
        
        # Provider info
        print("Press BACKSPACE to continue to provider info...")
        wait_for_backspace()
        
        provider_fields = [
            (row[9], 2),    # Provider + 2 tabs
            (row[10], 16),  # Physician + 16 tabs
            (row[11], 0),   # Diagnosis (no tab after)
        ]
        
        for value, tabs in provider_fields:
            type_safely(format_excel_value(value))
            pyautogui.press('tab', presses=tabs, interval=0.1)
        
        # Next patient or exit
        print(f"Finished patient {row[0]}. Press BACKSPACE for next or ESC to exit...")
        def on_press_continue(key):
            global should_continue
            if key == keyboard.Key.backspace:
                return False
            elif key == keyboard.Key.esc:
                should_continue = False
                return False
        with keyboard.Listener(on_press=on_press_continue) as listener:
            listener.join()

if __name__ == "__main__":
    print("Patient Data Entry Script (Excel Version)")
    print("----------------------------------------")
    print("Controls:")
    print("- BACKSPACE: Progress through each section")
    print("- ESC: Exit script anytime")
    print("\nMake sure:")
    print("1. Excel is closed before running")
    print("2. Cursor is in FIRST NAME field when starting")
    
    process_patients('patients.xlsx')  # Change to your filename
    
    print("\nCaptured Patient Codes:")
    for i, code in enumerate(patient_codes, 1):
        print(f"{i}. {code if code else 'Empty'}")
    
    print("\nScript completed.")
