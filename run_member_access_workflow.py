import subprocess
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def main():
    try:
        # Step 1: Open a file explorer dialog to select the input file
        print("Please select the input CSV file...")
        Tk().withdraw()  # Hide the root Tkinter window
        input_file = askopenfilename(
            title="Select Input CSV File",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )

        if not input_file:
            print("No file selected. Exiting workflow.")
            return

        print(f"Selected file: {input_file}")

        # Step 2: Run the script to create new user access codes
        print("Running create_new_user_access_codes.py...")
        subprocess.run(["python", "create_new_user_access_codes.py", input_file], check=True)
        print("Finished creating new user access codes.")

        # Step 3: Run the script to send emails
        print("Running outlook_templates.py...")
        subprocess.run(["python", "outlook_templates.py"], check=True)
        print("Finished sending emails.")

    except subprocess.CalledProcessError as e:
        print(f"An error occurred while running a script: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    main()