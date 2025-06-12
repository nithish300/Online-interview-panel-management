from flask import Flask, request, render_template
import pandas as pd
import os

app = Flask(__name__)
EXCEL_FILE = r"Panel.xlsx"
# Ensure the Excel file exists

if not os.path.exists(EXCEL_FILE):
    # Create an empty DataFrame with the required columns
    columns = [
        'Panel Email ID', 'Panel Grade', 'Panel Evaluation Round',
        'Panel Name', 'Panel Contact Number', 'TSR Code / Name',
        'Panel Account Name', 'Competency Code', 'Panel Work Geo',
        'Skills'  # Single column for all skills
    ]
    df = pd.DataFrame(columns=columns)
    df.to_excel(EXCEL_FILE, index=False)

# Flask app to create and save panel profiles

@app.route('/')
def index():
    return render_template('create_profile.html')

@app.route('/save', methods=['POST'])
def save_data():
    try:
        # Get form data
        email = request.form.get('email')
        grade = request.form.get('grade')
        evaluation_round = request.form.get('evaluation_round')
        name = request.form.get('name')
        contact = request.form.get('contact')
        tsr = request.form.get('tsr')
        account = request.form.get('account')
        competency = request.form.get('competency')
        geo = request.form.get('geo')

        # Collect all 8 skills into a list
        skills = [
            request.form.get(f'skill_{i}') for i in range(1, 9)
        ]
        # Filter out empty values
        skills = [skill for skill in skills if skill]

        # Convert to comma-separated string (or use json.dumps(skills) if preferred)
        skills_str = ", ".join(skills)

        # Prepare data dictionary with a single Skills column
        data = {
            'Panel Email ID': [email],
            'Panel Grade': [grade],
            'Panel Evaluation Round': [evaluation_round],
            'Panel Name': [name],
            'Panel Contact Number': [contact],
            'TSR Code / Name': [tsr],
            'Panel Account Name': [account],
            'Competency Code': [competency],
            'Panel Work Geo': [geo],
            'Skills': [skills_str]  # ← All skills stored here
        }
        # Ensure skills are stored as a single string in the DataFrame
        # Create DataFrame for new row
        df_new = pd.DataFrame(data)

        # If Excel exists, read and append; else create new
        if os.path.exists(EXCEL_FILE):
            df_existing = pd.read_excel(EXCEL_FILE)
            df = pd.concat([df_existing, df_new], ignore_index=True)
        else:
            df = df_new

        # Save back to Excel
        df.to_excel(EXCEL_FILE, index=False)

        return "Data saved successfully!"

    except Exception as e:
        return f"Error saving data: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True)
# Ensure the Flask app runs in debug mode for development
# and testing purposes. This allows for easier debugging and live reloading.
# The app will create an Excel file if it does not exist    