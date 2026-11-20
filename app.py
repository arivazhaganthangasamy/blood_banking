from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify
from werkzeug.security import generate_password_hash, check_password_hash
import pandas as pd
import os
from datetime import datetime, timedelta
import json

app = Flask(__name__)
app.secret_key = 'blooddonation_secret_2024'

EXCEL_FILE = 'blood_data.xlsx'

BLOOD_BANKS = [
    {"id": 1, "name": "Rotary TTK Blood Bank", "location": "Adyar, Chennai", "contact": "044-24918586", "lat": 13.0067, "lng": 80.2571, "blood_types": ["A+", "A-", "B+", "B-", "O+", "O-", "AB+", "AB-"]},
    {"id": 2, "name": "Voluntary Health Services Blood Bank", "location": "Adyar, Chennai", "contact": "044-22541965", "lat": 13.0012, "lng": 80.2565, "blood_types": ["A+", "B+", "O+", "AB+", "O-"]},
    {"id": 3, "name": "Government General Hospital Blood Bank", "location": "Park Town, Chennai", "contact": "044-25305000", "lat": 13.0827, "lng": 80.2707, "blood_types": ["A+", "A-", "B+", "B-", "O+", "O-", "AB+", "AB-"]},
    {"id": 4, "name": "Apollo Hospital Blood Bank", "location": "Greams Road, Chennai", "contact": "044-28293333", "lat": 13.0569, "lng": 80.2486, "blood_types": ["A+", "B+", "O+", "O-", "AB+"]},
    {"id": 5, "name": "Fortis Malar Hospital Blood Bank", "location": "Adyar, Chennai", "contact": "044-42892222", "lat": 13.0005, "lng": 80.2559, "blood_types": ["A+", "B+", "B-", "O+", "AB+", "AB-"]},
    {"id": 6, "name": "Sri Ramachandra Medical Centre Blood Bank", "location": "Porur, Chennai", "contact": "044-45928888", "lat": 13.0358, "lng": 80.1573, "blood_types": ["A+", "A-", "B+", "O+", "O-", "AB+"]},
    {"id": 7, "name": "Madras Medical Mission Blood Bank", "location": "Mogappair, Chennai", "contact": "044-26565961", "lat": 13.0839, "lng": 80.1746, "blood_types": ["A+", "B+", "O+", "O-"]},
    {"id": 8, "name": "Stanley Medical College Blood Bank", "location": "Royapuram, Chennai", "contact": "044-25281201", "lat": 13.1085, "lng": 80.2859, "blood_types": ["A+", "A-", "B+", "B-", "O+", "O-", "AB+", "AB-"]},
    {"id": 9, "name": "Kilpauk Medical College Blood Bank", "location": "Kilpauk, Chennai", "contact": "044-26421111", "lat": 13.0811, "lng": 80.2285, "blood_types": ["A+", "B+", "O+", "AB+"]},
    {"id": 10, "name": "JIPMER Chennai Blood Bank", "location": "Anna Nagar, Chennai", "contact": "044-26153215", "lat": 13.0845, "lng": 80.2101, "blood_types": ["A+", "A-", "O+", "O-", "B+"]},
]

HOSPITAL_BLOOD_DATA = {
    "A+": ["Rotary TTK", "VHS Blood Bank", "GGH Blood Bank", "Apollo", "SRMC"],
    "A-": ["Rotary TTK", "GGH Blood Bank", "SRMC", "Stanley Medical"],
    "B+": ["All Blood Banks"],
    "B-": ["Rotary TTK", "GGH Blood Bank", "Fortis Malar", "Stanley Medical"],
    "O+": ["All Blood Banks"],
    "O-": ["Rotary TTK", "VHS Blood Bank", "GGH Blood Bank", "Apollo", "SRMC", "MMM", "Stanley Medical"],
    "AB+": ["Rotary TTK", "GGH Blood Bank", "Apollo", "SRMC", "Fortis Malar", "Kilpauk Medical"],
    "AB-": ["Rotary TTK", "GGH Blood Bank", "Fortis Malar", "Stanley Medical"],
}

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            pd.DataFrame(columns=['id','username','email','password_hash','role','created_at']).to_excel(writer, sheet_name='Users', index=False)
            pd.DataFrame(columns=['id','donor_name','email','phone','blood_type','blood_bank_id','blood_bank_name','donation_date','next_eligible_date','status']).to_excel(writer, sheet_name='Donations', index=False)
            pd.DataFrame(columns=['id','requester_name','email','phone','blood_type','urgency','hospital','message','request_date','status']).to_excel(writer, sheet_name='BloodRequests', index=False)
            pd.DataFrame(columns=['id','user_email','blood_type','blood_bank_id','blood_bank_name','added_date']).to_excel(writer, sheet_name='Wishlist', index=False)
            pd.DataFrame(columns=['id','hospital_name','location','contact','blood_type','units_available','updated_at']).to_excel(writer, sheet_name='HospitalBlood', index=False)

def read_sheet(sheet_name):
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
    except:
        return pd.DataFrame()

def write_sheet(df, sheet_name):
    try:
        with pd.ExcelFile(EXCEL_FILE) as xls:
            sheets = {s: pd.read_excel(xls, sheet_name=s) for s in xls.sheet_names}
        sheets[sheet_name] = df
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            for name, data in sheets.items():
                data.to_excel(writer, sheet_name=name, index=False)
        return True
    except Exception as e:
        print(f"Error writing sheet: {e}")
        return False

init_excel()

# ─── AUTH ────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/signup', methods=['GET', 'POST'])
def signup():
    if request.method == 'POST':
        username = request.form['username']
        email = request.form['email']
        password = request.form['password']
        role = request.form.get('role', 'user')
        users = read_sheet('Users')
        if not users.empty and email in users['email'].values:
            flash('Email already registered!', 'error')
            return render_template('signup.html')
        new_id = 1 if users.empty else int(users['id'].max()) + 1
        new_user = pd.DataFrame([{'id': new_id, 'username': username, 'email': email,
                                   'password_hash': generate_password_hash(password),
                                   'role': role, 'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')}])
        updated = pd.concat([users, new_user], ignore_index=True)
        write_sheet(updated, 'Users')
        flash('Account created! Please login.', 'success')
        return redirect(url_for('login'))
    return render_template('signup.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form['email']
        password = request.form['password']
        users = read_sheet('Users')
        if not users.empty:
            user_row = users[users['email'] == email]
            if not user_row.empty:
                user = user_row.iloc[0]
                if check_password_hash(str(user['password_hash']), password):
                    session['user_id'] = int(user['id'])
                    session['username'] = str(user['username'])
                    session['email'] = str(user['email'])
                    session['role'] = str(user['role'])
                    if user['role'] == 'hospital':
                        return redirect(url_for('hospital_dashboard'))
                    return redirect(url_for('user_dashboard'))
        flash('Invalid email or password!', 'error')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('index'))

# ─── HOSPITAL ────────────────────────────────────────────────────────────────

@app.route('/hospital/dashboard')
def hospital_dashboard():
    if session.get('role') != 'hospital':
        return redirect(url_for('login'))
    blood_data = read_sheet('HospitalBlood')
    my_data = blood_data[blood_data['hospital_name'] == session['username']] if not blood_data.empty else pd.DataFrame()
    return render_template('hospital_dashboard.html', blood_data=my_data.to_dict('records'))

@app.route('/hospital/upload_blood', methods=['GET', 'POST'])
def upload_blood():
    if session.get('role') != 'hospital':
        return redirect(url_for('login'))
    if request.method == 'POST':
        blood_data = read_sheet('HospitalBlood')
        new_id = 1 if blood_data.empty else int(blood_data['id'].max()) + 1
        new_entry = pd.DataFrame([{
            'id': new_id,
            'hospital_name': session['username'],
            'location': request.form['location'],
            'contact': request.form['contact'],
            'blood_type': request.form['blood_type'],
            'units_available': request.form['units'],
            'updated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }])
        updated = pd.concat([blood_data, new_entry], ignore_index=True)
        write_sheet(updated, 'HospitalBlood')
        flash('Blood availability updated!', 'success')
        return redirect(url_for('hospital_dashboard'))
    return render_template('upload_blood.html')

# ─── USER ─────────────────────────────────────────────────────────────────────

@app.route('/user/dashboard')
def user_dashboard():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    return render_template('user_dashboard.html')

@app.route('/user/donate', methods=['GET', 'POST'])
def donate():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    if request.method == 'POST':
        email = session['email']
        donations = read_sheet('Donations')
        if not donations.empty:
            user_donations = donations[donations['email'] == email]
            if not user_donations.empty:
                last = pd.to_datetime(user_donations['donation_date'].max())
                next_eligible = last + timedelta(weeks=8)
                if datetime.now() < next_eligible:
                    flash(f'You can donate again after {next_eligible.strftime("%d %b %Y")} (8 weeks rule).', 'error')
                    return redirect(url_for('donate'))
        new_id = 1 if donations.empty else int(donations['id'].max()) + 1
        today = datetime.now()
        new_entry = pd.DataFrame([{
            'id': new_id,
            'donor_name': request.form['name'],
            'email': email,
            'phone': request.form['phone'],
            'blood_type': request.form['blood_type'],
            'blood_bank_id': request.form['blood_bank_id'],
            'blood_bank_name': request.form['blood_bank_name'],
            'donation_date': today.strftime('%Y-%m-%d'),
            'next_eligible_date': (today + timedelta(weeks=8)).strftime('%Y-%m-%d'),
            'status': 'Scheduled'
        }])
        updated = pd.concat([donations, new_entry], ignore_index=True)
        write_sheet(updated, 'Donations')
        flash('Donation registered successfully! Thank you for saving lives!', 'success')
        return redirect(url_for('user_dashboard'))
    return render_template('donate.html', blood_banks=BLOOD_BANKS)

@app.route('/user/need_blood', methods=['GET', 'POST'])
def need_blood():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    blood_type = request.args.get('blood_type', '')
    available_banks = []
    if blood_type:
        available_banks = [b for b in BLOOD_BANKS if blood_type in b['blood_types']]
    return render_template('need_blood.html', blood_banks=available_banks, blood_type=blood_type,
                           all_blood_types=["A+", "A-", "B+", "B-", "O+", "O-", "AB+", "AB-"])

@app.route('/user/wishlist/add', methods=['POST'])
def add_wishlist():
    if 'user_id' not in session:
        return jsonify({'success': False})
    data = request.json
    wishlist = read_sheet('Wishlist')
    new_id = 1 if wishlist.empty else int(wishlist['id'].max()) + 1
    new_entry = pd.DataFrame([{
        'id': new_id,
        'user_email': session['email'],
        'blood_type': data['blood_type'],
        'blood_bank_id': data['blood_bank_id'],
        'blood_bank_name': data['blood_bank_name'],
        'added_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }])
    updated = pd.concat([wishlist, new_entry], ignore_index=True)
    write_sheet(updated, 'Wishlist')
    return jsonify({'success': True, 'message': 'Added to wishlist!'})

@app.route('/user/wishlist')
def view_wishlist():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    wishlist = read_sheet('Wishlist')
    my_wishlist = []
    if not wishlist.empty:
        my_data = wishlist[wishlist['user_email'] == session['email']]
        my_wishlist = my_data.to_dict('records')
    return render_template('wishlist.html', wishlist=my_wishlist)

@app.route('/user/wishlist/remove/<int:item_id>')
def remove_wishlist(item_id):
    wishlist = read_sheet('Wishlist')
    if not wishlist.empty:
        wishlist = wishlist[wishlist['id'] != item_id]
        write_sheet(wishlist, 'Wishlist')
    flash('Removed from wishlist.', 'success')
    return redirect(url_for('view_wishlist'))

@app.route('/about')
def about():
    return render_template('about.html')

@app.route('/api/blood_banks')
def api_blood_banks():
    return jsonify(BLOOD_BANKS)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000,debug=True)
