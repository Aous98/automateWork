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
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.02)
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.02)
    try:
        code = pyperclip.paste()
        patient_codes.append(code)
        print(f"Captured patient code: {code}")
    except:
        print("Could not capture patient code")
        patient_codes.append("Unknown")

def format_excel_value(value):
    """Convert Excel values to proper strings"""
    if isinstance(value, datetime):
        return value.strftime("%m/%d/%Y")  # Format dates as MM/DD/YYYY
    elif value is None:
        return ""  # Handle empty cells
    else:
        return str(value)  # Convert everything else to string

def process_patients(file_path):
    global should_continue, patient_codes
    
    # Load Excel workbook
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    
    print("Position cursor in FIRST NAME field and press BACKSPACE to start...")
    wait_for_backspace()
    
    for row in sheet.iter_rows(min_row=2, values_only=True):  # Skip header
        if not should_continue:
            break
            
        # 1. Process first name to birthday
        # First name (full content)
        first_name = format_excel_value(row[0])
        #pyperclip.copy(first_name)
        #pyperclip.paste()
        pyautogui.write(first_name, interval=0.0)
        #time.sleep(0.05)

        pyautogui.press('tab')
        pyautogui.press('tab')
        # Last name (full content)
        last_name = format_excel_value(row[1])
        pyautogui.write(last_name)
        #time.sleep(0.05)
        pyautogui.press('tab')
        pyautogui.press('tab')
        # Patient code capture
        capture_patient_code()
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        # Address (full content)
        
        address = format_excel_value(row[2])
        pyautogui.write(address, interval=0.05)
        #time.sleep(0.05)
        pyautogui.press('tab')
        pyautogui.press('tab')
        
        # City (full content)
        city = format_excel_value(row[3])
        pyautogui.write(city, interval=0.0)
        #time.sleep(0.05)
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        
        # Zip (formatted)
        zip_code = format_excel_value(row[4])
        pyautogui.write(zip_code, interval=0.0)
        #time.sleep(0.05)
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        
        # Gender (full content)
        gender = format_excel_value(row[5])
        pyautogui.write(gender, interval=0.0)
        #time.sleep(0.05)
        pyautogui.press('tab')
        
        # Birthday (formatted date)
        #row[6] = str(row[6])
        birthday = format_excel_value(row[6])
        pyautogui.write(birthday, interval=0.0)
        
        # 2. Continue to insurance info
        print("Press BACKSPACE to continue to insurance info...")
        wait_for_backspace()
        
        # 3. Process insurance info
        # Payer
        payer = format_excel_value(row[7])
        pyautogui.write(payer, interval=0.0)
       # time.sleep(0.05)
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        
        # Member ID
        member_id = format_excel_value(row[8])
        pyautogui.write(member_id, interval=0.0)
        
        # 4. Continue to provider info
        print("Press BACKSPACE to continue to provider info...")
        wait_for_backspace()
        
        # 5. Process provider info
        # Provider
        provider = format_excel_value(row[9])
        pyautogui.write(provider, interval=0.0)
       # time.sleep(0.05)
        pyautogui.press('tab')
        pyautogui.press('tab')
        
        # Physician
        physician = format_excel_value(row[10])
        pyautogui.write(physician, interval=0.0)
       # time.sleep(0.05)
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
       # time.sleep(0.05)

        # Diagnosis
        diagnosis = format_excel_value(row[11])
        pyautogui.write(diagnosis, interval=0.0)
        
        # 6. Next patient or exit
        print(f"Finished patient {first_name} {last_name}. Press BACKSPACE for next or ESC to exit...")
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
    print("Note: Make sure Excel file has these columns in order:")
    print("First Name, Last Name, Address, City, Zip, Gender, Birthday, Payer, Member ID, Provider, Physician, Diagnosis")
    
    process_patients('TEMPLATE.xlsx')  # Change to your filename
    
    print("\nCaptured Patient Codes:")
    for i, code in enumerate(patient_codes, 1):
        print(code)
    
    print("\nScript completed. Press any key to exit...")
    #input()
