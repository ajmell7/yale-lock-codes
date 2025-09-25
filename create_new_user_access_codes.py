import os
import random
import pandas as pd
import pytz
import argparse
from datetime import datetime, timedelta
from dotenv import load_dotenv
from seam import Seam

# Load environment variables
load_dotenv()
SEAM_API_KEY = os.getenv("SEAM_API_KEY") 
seam = Seam(api_key=SEAM_API_KEY)

# Set timezone to Pacific Time
pacific = pytz.timezone("America/Los_Angeles")

# Generate new random 6-digit user access code
def generate_user_code():
    user_code = ""
    for _ in range(6):
        user_code += str(random.randint(0, 9))
    return user_code

# Create new users and assign access codes
def create_new_users(input_file):
    # Read the input CSV file
    df = pd.read_csv(input_file)

    # Get Yale lock device ID
    devices = seam.devices.list()
    device = next((d for d in devices if d.device_type == "yale_lock"), None)  # Remove once there is only a single device
    yale_lock_device_id = device.device_id if device else None

    if yale_lock_device_id:
        users_created = []
        for i in range(len(df)):
            user_name = df["First Name"][i] + " " + df["Last Name"][i]
            user_email = df["Email"][i]
            class_start_date_str = df["Class Start"][i]
            access_start_date = datetime.strptime(class_start_date_str, "%m/%d/%Y")
            access_start_date = pacific.localize(access_start_date) # Pacific Time zone
            class_end_date_str = df["Class End"][i]
            class_end_date = datetime.strptime(class_end_date_str, "%m/%d/%Y")
            class_end_date = pacific.localize(class_end_date)  # Pacific Time zone
            access_end_date = class_end_date + timedelta(days=7)
            user_code = generate_user_code()

            # Send new access code to Yale lock
            access_code = seam.access_codes.create(
                device_id=yale_lock_device_id,
                code=user_code,
                name=user_name,
                starts_at=access_start_date.isoformat(),
                ends_at=access_end_date.isoformat()
            )
            users_created.append([user_name, user_email, user_code, access_start_date, access_end_date])

            print(f"Created access code {access_code.code} for user {user_name}.")

        # Save the created users to a new CSV file
        users_df = pd.DataFrame(users_created, columns=["Name", "Email", "Access Code", "Access Start Date", "Access End Date"])
        users_df.to_csv("csv_files/users_created.csv", index=False)

def main():
    # Parse command-line arguments
    parser = argparse.ArgumentParser(description="Create new user access codes for Yale locks.")
    parser.add_argument("input_file", help="Path to the input CSV file containing user data.")
    args = parser.parse_args()

    # Call the create_new_users function with the input file
    create_new_users(args.input_file)

if __name__ == "__main__":
    main()