from flask import Flask, render_template, request
import random
import string
import os
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
import gspread
import json

app = Flask(__name__)

# Load environment variables from .env file
load_dotenv()

# Fetch the Google Sheets credentials JSON from environment variable
SERVICE_ACCOUNT_JSON = os.getenv('GOOGLE_APPLICATION_CREDENTIALS')

# Check if the environment variable is set
if SERVICE_ACCOUNT_JSON is None:
    print("Error: GOOGLE_APPLICATION_CREDENTIALS environment variable is not set.")
else:
    print("Environment variable for Google Sheets credentials is set.")

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
# Convert the string back to a dictionary
service_account_info = json.loads(SERVICE_ACCOUNT_JSON)
# Authenticate with Google Sheets
credentials = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
client = gspread.authorize(credentials)

# Open the Google Sheet
SPREADSHEET_ID = "1Yi-QemY97IEgQsWiOZWsmaz6zuwPr726iDmxCVzP70w"
print(f"Using Spreadsheet ID: {SPREADSHEET_ID}")
sheet = client.open_by_key(SPREADSHEET_ID).sheet1  # Access the first sheet


# Helper function to generate a unique ID
def generate_unique_id():
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=5))

@app.route("/", methods=["GET"])
def home():
    return render_template("index.html")

@app.route("/submit", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        try:    
        # Extract form data
            date = request.form["date"]
            name = request.form["name"]
            referrer_contact = request.form["referrer_contact"]
            case_from = request.form["case_from"]
            family_member_contact = request.form["family_member_contact"]
            family_member_name = request.form["family_member_name"]
            address1 = request.form["address1"]
            case_detail = request.form["case_detail"]


            # Generate a unique ID for the new row
            unique_id = generate_unique_id()

            # Insert the new data into the Excel sheet
            new_row = [
                unique_id, date, name, referrer_contact, case_from, family_member_contact,family_member_name,
                address1, case_detail
            ]
            sheet.append_row(new_row)

            return render_template("id_display.html", unique_id=unique_id)
        except Exception as e:
            # Log the error and show a user-friendly message
            print(f"Error in /submit: {str(e)}")
            return f"An error occurred while processing your request. Please try again later.Error :{str(e)}"
    return render_template("index.html")
@app.route("/referred_case_status", methods=["GET", "POST"])
def referred_case_status():
    case_data = None
    error_message = None
    
    if request.method == "POST":
        # Get the ID entered by the user
        case_id = request.form["case_id"]
        
        try:
            # Fetch all rows from the Google Sheet
            rows = sheet.get_all_values()


            # Search for the ID in the first column (Unique ID column)
            for row in rows:
                if row[0] == case_id:  # Assuming the Unique ID is in the first column
                    case_data = row
                    break

            # If no matching ID is found
            if not case_data:
                error_message = "No case found with this ID. Please check the ID and try again."
        except Exception as e:
            error_message = f"Error reading Google Sheet: {str(e)}"
        
    # Render the referred_case_status page with case_data (row) or error_message
    return render_template("referred_case_status.html", case_data=case_data, error_message=error_message)
    

if __name__ == "__main__":
    app.run(debug=True)
