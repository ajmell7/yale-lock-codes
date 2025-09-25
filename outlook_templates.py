import requests
import os
import pandas as pd

from datetime import datetime
from dotenv import load_dotenv

# load environment variables
load_dotenv()

# get secrets
TENANT_ID = os.getenv("TENANT_ID") 
CLIENT_ID = os.getenv("CLIENT_ID") 
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

if not TENANT_ID or not CLIENT_ID or not CLIENT_SECRET:
    raise ValueError("Missing one or more required environment variables: TENANT_ID, CLIENT_ID, CLIENT_SECRET")

def get_access_token(tenant_id, client_id, client_secret):
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default"
    }
    response = requests.post(url, data=data, timeout=10)
    if response.status_code == 200:
        return response.json()["access_token"]
    else:
        raise Exception(f"Failed to get access token: {response.status_code}, {response.text}")

def send_email_via_graph_api(access_token, user_name, subject, body, to_email, from_email):
    url = f"https://graph.microsoft.com/v1.0/users/{from_email}/sendMail"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    email_data = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": body
            },
            "toRecipients": [
                {"emailAddress": {"address": to_email}}
            ]
        }
    }
    response = requests.post(url, headers=headers, json=email_data)
    if response.status_code == 202:
        print(f"Email sent successfully to {user_name}!")
    else:
        print(f"Failed to send email to {user_name}: {response.status_code}, {response.text}")

def load_html_template(file_path):
    """Load the HTML template from a file."""
    with open(file_path, "r") as file:
        return file.read()

def generate_email_body(template, user_name, access_code, start_date, end_date):
    """Replace placeholders in the template with actual data."""
    return template.format(
        user_name=user_name,
        access_code=access_code,
        start_date=start_date,
        end_date=end_date
    )

def main():
     # Load the HTML template
    html_template = load_html_template("email_templates/class_email_template.html")

    # Load the CSV file
    df = pd.read_csv("csv_files/users_created.csv", dtype={"Access Code": str})

    current_access_token = get_access_token(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    
    for _, row in df.iterrows():
        user_name = row["Name"]
        user_email = row["Email"]
        access_code = row["Access Code"]
        # Parse and format the dates
        start_date = datetime.fromisoformat(row["Access Start Date"]).strftime("%m/%d/%y")
        end_date = datetime.fromisoformat(row["Access End Date"]).strftime("%m/%d/%y")
        # Generate the email body
        email_body = generate_email_body(
            html_template, user_name, access_code, start_date, end_date
        )
    
        send_email_via_graph_api(
            access_token=current_access_token,
            user_name=user_name,
            subject="TTP Door Access Code and Policies",
            body=email_body,
            to_email="ajmell7@gmail.com",
            from_email="hello@thirdplacepottery.com"
        )


if __name__ == "__main__":
    main()