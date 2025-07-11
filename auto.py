import csv
import pyautogui
import time
from pynput import keyboard

# Global variables
should_continue = True
patient_codes = []  # List to store captured patient codes

def wait_for_enter():
    """Wait for Enter key press to continue"""
    def on_press(key):
        if key == keyboard.Key.enter:
            return False  # Stop listener when Enter is pressed
    
    # Create listener
    with keyboard.Listener(on_press=on_press) as listener:
        listener.join()

def capture_patient_code():
    """Copy the current cell content (patient code) and save to list"""
    # Select all text in current cell
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.1)
    # Copy the text

    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.1)
    # Get from clipboard
    try:
        import pyperclip
        code = pyperclip.paste()
        patient_codes.append(code)
        print(f"Captured patient code: {code}")
    except:
        print("Could not capture patient code")
        patient_codes.append("Unknown")

def process_patients(file_path):
    global should_continue, patient_codes
    
    with open(file_path, 'r') as file:
        reader = csv.reader(file)
        next(reader, None)  # Skip header
        
        print("Position cursor in FIRST NAME field and press ENTER to start...")
        wait_for_enter()  # Wait for initial Enter press
        #time.sleep(0.5)  # Small delay before starting
        
        for row in reader:
            if not should_continue:
                break
                
            # 1. Process first name to birthday (first 7 fields)
           
            ## First name
            pyautogui.write(row[0])
            pyautogui.press('tab',presses=2)
            
            
            ## Last name
            pyautogui.write(row[1])
            pyautogui.press('tab',presses=2)
            
            ## Here it should be patient code
            #time.sleep(2)

            capture_patient_code()  # This will save the patient code
            pyautogui.press('tab',presses=3)
          
            # Address
            pyautogui.write(row[2])
            pyautogui.press('tab',presses=2)

            
            # City
            pyautogui.write(row[3])
            pyautogui.press('tab',presses=3)

            
            ## Zip
            pyautogui.write(row[4])
            pyautogui.press('tab',presses=11)

            
            ## Gender
            pyautogui.write(row[5])
            pyautogui.press('tab')
            
            ## Birthday
            pyautogui.write(row[6])

            # 2. Wait for Enter press to continue
            print("Press ENTER to continue to insurance info...")
            wait_for_enter()
            
            # 3. Process insurance info (next 2 fields)
            ## Payer
            pyautogui.write(row[7])
            pyautogui.press('tab',presses=4)


            ## Member ID
            pyautogui.write(row[8])
         
            # 4. Wait for Enter press to continue
            print("Press ENTER to continue to provider info...")
            wait_for_enter()
            
            # 5. Process provider info (last 3 fields)
            ## Provider
            pyautogui.write(row[9])
            pyautogui.press('tab',presses=2)

            ## Doctor
            pyautogui.write(row[10])
            pyautogui.press('tab',presses=16)

            
            ## Diagnose
            pyautogui.write(row[11])

            # 6. Wait for Enter before next patient
            print(f"Finished patient {row[0]} {row[1]}. Press ENTER for next patient or ESC to exit...")
            
            def on_press_continue(key):
                global should_continue
                if key == keyboard.Key.enter:
                    return False  # Continue
                elif key == keyboard.Key.esc:
                    should_continue = False
                    return False  # Exit
            
            with keyboard.Listener(on_press=on_press_continue) as listener:
                listener.join()

if __name__ == "__main__":
    print("Patient Data Entry Script")
    print("------------------------")
    print("Controls:")
    print("- ENTER: Start script and continue through all steps")
    print("- ESC: Exit script anytime")
    print("\nMake sure your cursor is in the first name field before starting!")
    
    process_patients('bulls.csv')
    
    # Print all captured patient codes at the end
    print("\nCaptured Patient Codes:")
    for i, code in enumerate(patient_codes, 1):
        print(f"{i}. {code}")
    
    print("Script completed or terminated.")