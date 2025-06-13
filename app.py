from flask import Flask, request, render_template, jsonify
import pandas as pd
import os
import traceback

app = Flask(__name__)
EXCEL_FILE = "data/Panel.xlsx"
EXCEL_SLOT_FILE = "data/Panel_slot.xlsx"


# Create Excel if not exists
if not os.path.exists(EXCEL_FILE):
    columns = [
        'Panel Email ID', 'Panel Grade', 'Panel Evaluation Round',
        'Panel Name', 'Panel Contact Number', 'TSR Code / Name',
        'Panel Account Name', 'Competency Code', 'Panel Work Geo', 'Skills'
    ]
    df = pd.DataFrame(columns=columns)
    df.to_excel(EXCEL_FILE, index=False)

if not os.path.exists(EXCEL_SLOT_FILE):
    slot_columns = [
        'Panel Email ID', 'Name', 'Competency', 'Geo',
        'Slot Start Date', 'Slot End Date', 'Slot Start Time', 'Number of Slots'
    ]
    df = pd.DataFrame(columns=slot_columns)
    df.to_excel(EXCEL_SLOT_FILE, index=False)



@app.route('/')
def index():
    return render_template('main.html')

@app.route('/get_profile', methods=['GET'])
def get_profile():
    email = request.args.get('email')
    if not email:
        return jsonify({"error": "Email is required"}), 400

    try:
        df = pd.read_excel(EXCEL_FILE)
        if email not in df['Panel Email ID'].values:
            return jsonify({"error": "Email not found"}), 404

        row = df[df['Panel Email ID'] == email].iloc[0]
        skills = row['Skills'].split(", ") if isinstance(row['Skills'], str) else []

        return jsonify({
            "email": row['Panel Email ID'],
            "grade": row['Panel Grade'],
            "evaluation_round": row['Panel Evaluation Round'],
            "name": row['Panel Name'],
            "contact": row['Panel Contact Number'],
            "tsr": row['TSR Code / Name'],
            "account": row['Panel Account Name'],
            "competency": row['Competency Code'],
            "geo": row['Panel Work Geo'],
            "skills": skills
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/save', methods=['POST'])
def save_data():
    try:
        email = request.form.get('email')
        grade = request.form.get('grade')
        evaluation_round = request.form.get('evaluation_round')
        name = request.form.get('name')
        contact = request.form.get('contact')
        tsr = request.form.get('tsr')
        account = request.form.get('account')
        competency = request.form.get('competency')
        geo = request.form.get('geo')

        skills = [request.form.get(f'skill_{i}') for i in range(1, 9)]
        skills = [s for s in skills if s]
        skills_str = ", ".join(skills)

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
            'Skills': [skills_str]
        }

        df_new = pd.DataFrame(data)

        if os.path.exists(EXCEL_FILE):
            df_existing = pd.read_excel(EXCEL_FILE)
            df = pd.concat([df_existing, df_new], ignore_index=True)
        else:
            df = df_new

        df.to_excel(EXCEL_FILE, index=False)
        return "Data saved successfully!"

    except Exception as e:
        return f"Error saving data: {str(e)}", 500


# 🟡 NEW ROUTE - Show Modify Form
@app.route('/modify_profile')
def modify_profile():
    return render_template('modify_profile.html')  # Make sure this matches the filename exactly

@app.route('/list_users')
def list_users():
    df = pd.read_excel(EXCEL_FILE)
    return df.to_html(index=False)


@app.route('/update_profile', methods=['POST'])
def update_profile():
    if not request.form:
        return "No data provided", 400

    # Get form data
    email = request.form.get('email')
    grade = request.form.get('grade', '').strip()
    evaluation_round = request.form.get('evaluation_round', '').strip()
    name = request.form.get('name', '').strip()
    contact = request.form.get('contact', '').strip()
    tsr = request.form.get('tsr', '').strip()
    account = request.form.get('account', '').strip()
    competency = request.form.get('competency', '').strip()
    geo = request.form.get('geo', '').strip()

    # Handle multiple skill fields
    skills = [request.form.get(f'skill_{i}', '').strip() for i in range(1, 9)]
    skills = [s for s in skills if s]  # Remove empty strings
    skills_str = ", ".join(skills)

    # Read Excel file
    df = pd.read_excel(EXCEL_FILE)

    # Rename columns temporarily for easier access
    df.rename(columns={
        'Panel Email ID': 'Email',
        'Panel Grade': 'Grade',
        'Panel Evaluation Round': 'Evaluation_Round',
        'Panel Name': 'Name',
        'Panel Contact Number': 'Contact',
        'TSR Code / Name': 'TSR',
        'Panel Account Name': 'Account',
        'Competency Code': 'Competency',
        'Panel Work Geo': 'Geo',
        'Skills': 'Skills'
    }, inplace=True)

    # Check if email exists
    if email not in df['Email'].values:
        new_row = {
            'Email': email,
            'Grade': grade,
            'Evaluation_Round': evaluation_round,
            'Name': name,
            'Contact': contact,
            'TSR': tsr,
            'Account': account,
            'Competency': competency,
            'Geo': geo,
            'Skills': skills_str
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        message = f"New profile for {email} added!"
    else:
        if grade:
            df.loc[df['Email'] == email, 'Grade'] = grade
        if evaluation_round:
            df.loc[df['Email'] == email, 'Evaluation_Round'] = evaluation_round
        if name:
            df.loc[df['Email'] == email, 'Name'] = name
        if contact:
            df.loc[df['Email'] == email, 'Contact'] = contact
        if tsr:
            df.loc[df['Email'] == email, 'TSR'] = tsr
        if account:
            df.loc[df['Email'] == email, 'Account'] = account
        if competency:
            df.loc[df['Email'] == email, 'Competency'] = competency
        if geo:
            df.loc[df['Email'] == email, 'Geo'] = geo
        df.loc[df['Email'] == email, 'Skills'] = skills_str
        message = f"Profile for {email} updated successfully!"

    # Rename back to original column names
    df.rename(columns={
        'Email': 'Panel Email ID',
        'Grade': 'Panel Grade',
        'Evaluation_Round': 'Panel Evaluation Round',
        'Name': 'Panel Name',
        'Contact': 'Panel Contact Number',
        'TSR': 'TSR Code / Name',
        'Account': 'Panel Account Name',
        'Competency': 'Competency Code',
        'Geo': 'Panel Work Geo',
        'Skills': 'Skills'
    }, inplace=True)

    # Save changes back to Excel
    df.to_excel(EXCEL_FILE, index=False)

    return message

import traceback

@app.route('/save_slot', methods=['POST'])
def save_slot():
    try:
        print("Form Data:", request.form)

        panel_email_id = request.form.get('panel_email_id')
        name = request.form.get('name')
        competency = request.form.get('competency')
        geo = request.form.get('geo')
        slot_start_date = request.form.get('slot_start_date')
        slot_end_date = request.form.get('slot_end_date')
        slot_time = request.form.get('slot_time')
        num_slots = request.form.get('num_slots')
        
        if not all([panel_email_id, name, competency, geo, slot_start_date, slot_end_date, slot_time, num_slots]):
            return "All fields are required.", 400

        if not num_slots.isdigit() or int(num_slots) <= 0:
            return "Number of slots must be a positive integer.", 400

        num_slots = int(num_slots)

        data = {
            'Panel Email ID': [panel_email_id],
            'Name': [name],
            'Competency': [competency],
            'Geo': [geo],
            'Slot Start Date': [slot_start_date],
            'Slot End Date': [slot_end_date],
            'Slot Start Time': [slot_time],
            'Number of Slots': [num_slots]
        }

        df_new = pd.DataFrame(data)

        if os.path.exists(EXCEL_SLOT_FILE):
            df_existing = pd.read_excel(EXCEL_SLOT_FILE)
            df = pd.concat([df_existing, df_new], ignore_index=True)
        else:
            df = df_new

        df.to_excel(EXCEL_SLOT_FILE, index=False)

        return "Slot data saved successfully!"

    except Exception as e:
        traceback.print_exc()  # Show detailed error
        print("Error:", str(e))
        return f"Error saving slot data: {str(e)}", 500
    
@app.route('/update_slot', methods=['POST'])
def update_slot():
    try:
        print("Form Data:", request.form)

        panel_email_id = request.form.get('email')
        name = request.form.get('name')
        competency = request.form.get('competency')
        geo = request.form.get('geo')
        slot_start_date = request.form.get('slot_start_date')
        slot_end_date = request.form.get('slot_end_date')
        slot_time = request.form.get('slot_time')
        num_slots = request.form.get('num_slots')

        if not all([panel_email_id, name, competency, geo, slot_start_date, slot_end_date, slot_time, num_slots]):
            print("Missing Fields")
            return "All fields are required.", 400

        if not num_slots.isdigit() or int(num_slots) <= 0:
            return "Number of slots must be a positive integer.", 400

        num_slots = int(num_slots)

        df = pd.read_excel(EXCEL_SLOT_FILE)

        expected_columns = [
            'Panel Email ID', 'Name', 'Competency', 'Geo',
            'Slot Start Date', 'Slot End Date', 'Slot Start Time', 'Number of Slots'
        ]
        for col in expected_columns:
            if col not in df.columns:
                return f"Required column '{col}' missing from {EXCEL_SLOT_FILE}", 500

        # Check if the email exists
        if panel_email_id not in df['Panel Email ID'].values:
            # Add new row
            new_row = {
                'Panel Email ID': panel_email_id,
                'Name': name,
                'Competency': competency,
                'Geo': geo,
                'Slot Start Date': slot_start_date,
                'Slot End Date': slot_end_date,
                'Slot Start Time': slot_time,
                'Number of Slots': num_slots
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            message = f"New slot for {panel_email_id} added!"
        else:
            # Update existing row
            df.loc[df['Panel Email ID'] == panel_email_id, 'Name'] = name
            df.loc[df['Panel Email ID'] == panel_email_id, 'Competency'] = competency
            df.loc[df['Panel Email ID'] == panel_email_id, 'Geo'] = geo
            df.loc[df['Panel Email ID'] == panel_email_id, 'Slot Start Date'] = slot_start_date
            df.loc[df['Panel Email ID'] == panel_email_id, 'Slot End Date'] = slot_end_date
            df.loc[df['Panel Email ID'] == panel_email_id, 'Slot Start Time'] = slot_time
            df.loc[df['Panel Email ID'] == panel_email_id, 'Number of Slots'] = num_slots
            message = f"Slot for {panel_email_id} updated successfully!"

        # Save back to Excel
        df.to_excel(EXCEL_SLOT_FILE, index=False)

        return message

    except Exception as e:
        import traceback
        traceback.print_exc()
        print("Error:", str(e))
        return f"Error updating slot: {str(e)}", 500
    

if __name__ == '__main__':
    app.run(debug=True)