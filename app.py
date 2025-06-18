from flask import Flask, request, render_template, jsonify, redirect, url_for, session, flash
import pandas as pd
import os
import traceback
import numpy as np
from flask.json.provider import DefaultJSONProvider

# Ensure the necessary directories exist
app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Change this in production

EXCEL_FILE = "data/Panel.xlsx"
LOGIN_FILE = "data/login.xlsx"

os.makedirs('data', exist_ok=True)

class CustomJSONEncoder(DefaultJSONProvider):
    def default(self, obj):
        if isinstance(obj, (np.integer, np.int64)):
            return int(obj)
        elif isinstance(obj, (np.floating, np.float64)):
            return float(obj)
        elif isinstance(obj, (pd.Timestamp, pd.Timedelta)):
            return str(obj)
        elif isinstance(obj, pd.Series):
            return obj.to_dict()
        return super().default(obj)

app.json = CustomJSONEncoder(app)

# Load login credentials from Excel
def load_login_credentials():
    if not os.path.exists(LOGIN_FILE):
        return {}
    df = pd.read_excel(LOGIN_FILE)
    credentials = df.set_index('Email').to_dict(orient='index')
    return {email: {'password': creds['Password'], 'role': creds['Role']} for email, creds in credentials.items()}

@app.route('/')
def index():
    if 'user_email' in session:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form.get('email')
        password = request.form.get('password')
        credentials = load_login_credentials()
        if email in credentials and credentials[email]['password'] == password:
            session['user_email'] = email
            session['user_role'] = credentials[email]['role']
            return redirect(url_for('dashboard'))
        else:
            flash("Invalid credentials")
            return redirect(url_for('login'))
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.pop('user_email', None)
    session.pop('user_role', None)
    return redirect(url_for('login'))

@app.route('/dashboard')
def dashboard():
    if 'user_email' not in session:
        return redirect(url_for('login'))
    role = session.get('user_role')
    if role == 'Admin':
        return render_template('main_admin.html')
    elif role == 'Employee':
        return render_template('main_employee.html')
    else:
        return render_template('unauthorized.html'), 403

def is_authorized():
    return 'user_email' in session and session['user_role'] in ['Admin', 'Employee']

@app.route('/create_profile')
def create_profile():
    if not is_authorized():
        return render_template('unauthorized.html'), 403
    return render_template('create_profile.html')

@app.route('/modify_profile')
def modify_profile():
    if not is_authorized():
        return render_template('unauthorized.html'), 403
    return render_template('modify_profile.html')

@app.route('/create_slot')
def create_slot():
    if not is_authorized():
        return render_template('unauthorized.html'), 403
    return render_template('create_slot.html')

@app.route('/modify_slot')
def modify_slot():
    if not is_authorized():
        return render_template('unauthorized.html'), 403
    return render_template('modify_slot.html')

@app.route('/manage_profile_admin')
def manage_profile_admin():
    return render_template('manage_profile_admin.html')

@app.route('/manage_slot_admin')
def manage_slot_admin():
    return render_template('manage_slot_admin.html')

@app.route('/import_profiles')
def import_profiles():
    if session.get('user_role') != 'Admin':
        return render_template('unauthorized.html'), 403
    return render_template('import_profiles.html')

@app.route('/view_interviewers')
def view_interviewers():
    if session.get('user_role') != 'Admin':
        return render_template('unauthorized.html'), 403

    try:
        df = pd.read_excel(EXCEL_FILE)
        
        # Ensure required columns exist
        expected_columns = ['Emp ID', 'Panel Name', 'Competency Code', 'Panel Work Geo', 'TSR Code / Name', 'Skills', 
                            'Slot Start Date', 'Slot End Date', 'Slot Start Time', 'Number of Slots']
        missing_cols = [col for col in expected_columns if col not in df.columns]
        if missing_cols:
            flash(f"Missing columns in Excel: {missing_cols}")
            return redirect(url_for('dashboard'))

        # Convert DataFrame to list of dictionaries
        interviewers = []

        for _, row in df.iterrows():
            start_date = str(row['Slot Start Date']) if pd.notna(row['Slot Start Date']) else None
            end_date = str(row['Slot End Date']) if pd.notna(row['Slot End Date']) else None
            start_time = str(row['Slot Start Time']) if pd.notna(row['Slot Start Time']) else None
            num_slots = int(row['Number of Slots']) if pd.notna(row['Number of Slots']) else 0
            
            # For now, assume all slots are available
            available_slots = num_slots  # You can reduce this as bookings come in
            
            interviewer_data = {
                "emp_id": row['Emp ID'],
                "name": row['Panel Name'],
                "competency": row['Competency Code'],
                "geo": row['Panel Work Geo'],
                "tsr": row['TSR Code / Name'],
                "skills": row['Skills'],
                "start_date": start_date,
                "end_date": end_date,
                "start_time": start_time,
                "num_slots": num_slots,
                "available_slots": available_slots
            }
            
            interviewers.append(interviewer_data)

        return render_template('view_interviewers.html', interviewers=interviewers)

    except Exception as e:
        print("Error loading interviewers:", str(e))
        return f"Error loading interviewers: {str(e)}", 500

# Ensure Panel.xlsx has all required columns
def ensure_columns(df):
    expected_columns = [
        'Emp ID', 'Panel Grade', 'Panel Evaluation Round',
        'Panel Name', 'Panel Contact Number', 'TSR Code / Name',
        'Panel Account Name', 'Competency Code', 'Panel Work Geo', 'Skills',
        'Slot Start Date', 'Slot End Date', 'Slot Start Time', 'Number of Slots'
    ]
    for col in expected_columns:
        if col not in df.columns:
            df[col] = ""
    return df

# Create empty file if not exists
if not os.path.exists(EXCEL_FILE):
    columns = [
        'Emp ID', 'Panel Grade', 'Panel Evaluation Round',
        'Panel Name', 'Panel Contact Number', 'TSR Code / Name',
        'Panel Account Name', 'Competency Code', 'Panel Work Geo', 'Skills',
        'Slot Start Date', 'Slot End Date', 'Slot Start Time', 'Number of Slots'
    ]
    df = pd.DataFrame(columns=columns)
    df.to_excel(EXCEL_FILE, index=False)
else:
    df = pd.read_excel(EXCEL_FILE)
    if 'Panel Email ID' in df.columns:
        df.rename(columns={'Panel Email ID': 'Emp ID'}, inplace=True)
    df.to_excel(EXCEL_FILE, index=False)


@app.route('/save_data', methods=['POST'])
def save_data():
    try:
        print("Received form:", request.form)
        emp_id = request.form.get('emp_id')
        grade = request.form.get('grade')
        evaluation_round = request.form.get('evaluation_round')
        name = request.form.get('name')
        contact = request.form.get('contact')
        tsr = request.form.get('tsr')
        account = request.form.get('account')
        competency = request.form.get('competency')
        geo = request.form.get('geo')
        skills = [request.form.get(f'skill_{i}') for i in range(1, 9)]
        skills = [s.strip() for s in skills if s]
        skills_str = ", ".join(skills)

        df = pd.read_excel(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else pd.DataFrame()

        if 'Emp ID' not in df.columns:
            df['Emp ID'] = ""

        if emp_id in df['Emp ID'].values:
            return f"Error: Emp ID '{emp_id}' already exists.", 400

        new_row = {
            'Emp ID': [emp_id],
            'Panel Grade': [grade],
            'Panel Evaluation Round': [evaluation_round],
            'Panel Name': [name],
            'Panel Contact Number': [contact],
            'TSR Code / Name': [tsr],
            'Panel Account Name': [account],
            'Competency Code': [competency],
            'Panel Work Geo': [geo],
            'Skills': [skills_str]
        }

        df_new = pd.DataFrame(new_row)
        df = pd.concat([df, df_new], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        return "Profile created successfully!"

    except Exception as e:
        print("Error saving data:", str(e))
        return f"Error saving data: {str(e)}", 500

@app.route('/update_profile', methods=['POST'])
def update_profile():
    try:
        emp_id = request.form.get('emp_id')
        grade = request.form.get('grade', '').strip()
        evaluation_round = request.form.get('evaluation_round', '').strip()
        name = request.form.get('name', '').strip()
        contact = request.form.get('contact', '').strip()
        tsr = request.form.get('tsr', '').strip()
        account = request.form.get('account', '').strip()
        competency = request.form.get('competency', '').strip()
        geo = request.form.get('geo', '').strip()

        # Collect skills (max 8)
        skills = [request.form.get(f'skill_{i}', '').strip() for i in range(1, 9)]
        skills = [s for s in skills if s]
        skills_str = ", ".join(skills)

        # Load Excel data
        df = pd.read_excel(EXCEL_FILE)
        df = ensure_columns(df)

        # Ensure Emp ID is treated as string
        if str(emp_id) not in df['Emp ID'].astype(str).values:
            return f"Emp ID '{emp_id}' not found!", 404

        # Update profile using string comparison
        df.loc[df['Emp ID'].astype(str) == str(emp_id), 'Panel Grade'] = grade
        df.loc[df['Emp ID'].astype(str) == str(emp_id), 'Panel Evaluation Round'] = evaluation_round
        df.loc[df['Emp ID'].astype(str) == str(emp_id), 'Panel Name'] = name
        df.loc[df['Emp ID'].astype(str) == str(emp_id), 'Panel Contact Number'] = contact
        df.loc[df['Emp ID'].astype(str) == str(emp_id), 'TSR Code / Name'] = tsr
        df.loc[df['Emp ID'].astype(str) == str(emp_id), 'Panel Account Name'] = account
        df.loc[df['Emp ID'].astype(str) == str(emp_id), 'Competency Code'] = competency
        df.loc[df['Emp ID'].astype(str) == str(emp_id), 'Panel Work Geo'] = geo
        df.loc[df['Emp ID'].astype(str) == str(emp_id), 'Skills'] = skills_str

        # Save changes back to Excel
        df.to_excel(EXCEL_FILE, index=False)

        return f"Profile for {emp_id} updated successfully!"

    except Exception as e:
        traceback.print_exc()
        print("Error updating profile:", str(e))
        return f"Error updating profile: {str(e)}", 500   

@app.route('/save_slot', methods=['POST'])
def save_slot():
    try:
        emp_id = request.form.get('emp_id')
        name = request.form.get('name')
        competency = request.form.get('competency')
        geo = request.form.get('geo')
        slot_start_date = request.form.get('slot_start_date')
        slot_end_date = request.form.get('slot_end_date')
        slot_time = request.form.get('slot_time')
        num_slots = request.form.get('num_slots')

        if not all([emp_id, name, competency, geo, slot_start_date, slot_end_date, slot_time, num_slots]):
            return "All fields are required.", 400

        if not num_slots.isdigit() or int(num_slots) <= 0:
            return "Number of slots must be a positive integer.", 400

        num_slots = int(num_slots)
        df = pd.read_excel(EXCEL_FILE)
        if str(emp_id) not in df['Emp ID'].astype(str).values:
            return f"Emp ID '{emp_id}' not found!", 404
        

        # Ensure all required columns exist
        df = ensure_columns(df) 
        df.loc[df['Emp ID'] == emp_id, 'Slot Start Date'] = slot_start_date
        df.loc[df['Emp ID'] == emp_id, 'Slot End Date'] = slot_end_date
        df.loc[df['Emp ID'] == emp_id, 'Slot Start Time'] = slot_time
        df.loc[df['Emp ID'] == emp_id, 'Number of Slots'] = num_slots

        df.to_excel(EXCEL_FILE, index=False)
        return f"Slot for {emp_id} saved successfully!"
    
    except Exception as e:
        print("Error saving slot:", str(e))
        return f"Error saving slot: {str(e)}", 500



@app.route('/update_slot', methods=['POST'])
def update_slot():
    try:
        emp_id = request.form.get('emp_id')
        name = request.form.get('name')
        competency = request.form.get('competency')
        geo = request.form.get('geo')
        slot_start_date = request.form.get('slot_start_date')
        slot_end_date = request.form.get('slot_end_date')
        slot_time = request.form.get('slot_time')
        num_slots = request.form.get('num_slots')

        if not all([slot_start_date, slot_end_date, slot_time, num_slots]):
            return "All fields are required!!", 400

        if not num_slots.isdigit() or int(num_slots) <= 0:
            return "Number of slots must be a positive integer.", 400

        num_slots = int(num_slots)

        df = pd.read_excel(EXCEL_FILE)
        df = ensure_columns(df)  # Make sure all expected columns exist

        # Use string-based matching for Emp ID
        if str(emp_id) not in df['Emp ID'].astype(str).values:
            return f"Emp ID '{emp_id}' not found!", 404

        # Update DataFrame using string-safe comparison
        df.loc[df['Emp ID'].astype(str) == str(emp_id), 'Slot Start Date'] = slot_start_date
        df.loc[df['Emp ID'].astype(str) == str(emp_id), 'Slot End Date'] = slot_end_date
        df.loc[df['Emp ID'].astype(str) == str(emp_id), 'Slot Start Time'] = slot_time
        df.loc[df['Emp ID'].astype(str) == str(emp_id), 'Number of Slots'] = num_slots

        # Save changes back to Excel
        df.to_excel(EXCEL_FILE, index=False)

        return f"Slot for {emp_id} updated successfully!"

    except Exception as e:
        print("Error updating slot:", str(e))
        return f"Error updating slot: {str(e)}", 500
    

@app.route("/save_bulk", methods=['POST'])
def save_bulk():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    if not file.filename.endswith('.xlsx'):
        return jsonify({"error": "Invalid file type. Only .xlsx allowed."}), 400

    temp_path = os.path.join("temp", file.filename)
    os.makedirs("temp", exist_ok=True)
    file.save(temp_path)

    try:
        merge_excels(EXCEL_FILE, temp_path)
        return jsonify({
            "message": "File processed and merged successfully",
        })
    except Exception as e:
        return jsonify({"error": f"Error merging files: {str(e)}"}), 500

def merge_excels(local_path, external_path, id_column='Emp ID'):
    """
    Merges two Excel files using a unique column (e.g., Emp ID)
    Keeps only unique rows based on id_column
    """
    try:
        # Use engine='openpyxl' for .xlsx files
        df_local = pd.read_excel(local_path, engine='openpyxl')
        df_external = pd.read_excel(external_path, engine='openpyxl')

        # Ensure the id_column exists
        if id_column not in df_local.columns or id_column not in df_external.columns:
            raise ValueError(f"Column '{id_column}' not found in one of the files")

        # Remove duplicates in both DataFrames
        df_external = df_external.drop_duplicates(subset=[id_column])
        df_local = df_local.drop_duplicates(subset=[id_column])

        # Merge and keep latest values from external
        df_combined = pd.concat([df_local, df_external]).drop_duplicates(subset=[id_column], keep='last')

        # Save merged data back to local file
        df_combined.to_excel(local_path, index=False)

        # Optional: delete temp file after merge
        if os.path.exists(external_path):
            os.remove(external_path)

        return "Merge successful"

    except Exception as e:
        print("Error during merge:", str(e))
        return f"Error merging files: {str(e)}"

@app.route("/search_employee", methods=['GET'])
def search_employee():
    employee_id = request.args.get('employee_id')
    if not employee_id:
        return jsonify({"error": "Employee ID is required"}), 400
    try:
        df = pd.read_excel(EXCEL_FILE)

        if 'Emp ID' not in df.columns:
            return jsonify({"error": "Emp ID column missing in file"}), 500

        # Compare as strings
        if str(employee_id) not in df['Emp ID'].astype(str).values:
            return jsonify({"error": "Employee ID not found"}), 404

        row = df[df['Emp ID'].astype(str) == str(employee_id)].iloc[0]

        def convert_value(val):
            if hasattr(val, 'item'):
                return val.item()
            elif isinstance(val, (pd.Timestamp, pd.Timedelta)):
                return str(val)
            else:
                return str(val) if not isinstance(val, str) and val is not None else val

        skills = row['Skills'].split(", ") if isinstance(row['Skills'], str) else []

        profile_data = {
            "emp_id": convert_value(row['Emp ID']),
            "grade": convert_value(row['Panel Grade']),
            "evaluation_round": convert_value(row['Panel Evaluation Round']),
            "name": convert_value(row['Panel Name']),
            "contact": convert_value(row['Panel Contact Number']),
            "tsr": convert_value(row['TSR Code / Name']),
            "account": convert_value(row['Panel Account Name']),
            "competency": convert_value(row['Competency Code']),
            "geo": convert_value(row['Panel Work Geo']),
            "skills": [convert_value(skill) for skill in skills],
        }

        return jsonify(profile_data)

    except Exception as e:
        print("Error during search:", str(e))
        return jsonify({"error": f"Error fetching profile: {str(e)}"}), 500
 
# Search employee endpoint

@app.route("/search_slot", methods=['GET'])
def search_slot():
    employee_id = request.args.get('employee_id')
    if not employee_id:
        return jsonify({"error": "Employee ID is required"}), 400

    try:
        df = pd.read_excel(EXCEL_FILE)
        if 'Emp ID' not in df.columns:
            return jsonify({"error": "Emp ID column missing in file"}), 500

        if str(employee_id) not in df['Emp ID'].astype(str).values:
            return jsonify({"error": "Employee ID not found"}), 404

        row = df[df['Emp ID'].astype(str) == str(employee_id)].iloc[0]

        def convert_value(val):
            if hasattr(val, 'item'):
                return val.item()
            elif isinstance(val, (pd.Timestamp, pd.Timedelta)):
                return str(val)
            else:
                return str(val) if not isinstance(val, str) and val is not None else val

        profile_data = {
            "start-date": convert_value(row['Slot Start Date']),
            "end-date": convert_value(row['Slot End Date']),
            "slot-time": convert_value(row['Slot Start Time']),
            "num-slots": convert_value(row['Number of Slots']),
            "name": convert_value(row['Panel Name']),
            "competency": convert_value(row['Competency Code']),
            "geo": convert_value(row['Panel Work Geo']),
        }

        return jsonify(profile_data)

    except Exception as e:
        print("Error during search:", str(e))
        return jsonify({"error": f"Error fetching profile: {str(e)}"}), 500   


if __name__ == '__main__':
    app.run(debug=True)
