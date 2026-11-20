# LifeStream - Blood Donation Management System

## Project Structure
```
blood_donation/
├── app.py                   # Main Flask application
├── requirements.txt         # Python dependencies
├── blood_data.xlsx          # Auto-created Excel database
└── templates/
    ├── base.html            # Base layout with nav, particles, animations
    ├── index.html           # Home page with hero, stats, blood type search
    ├── login.html           # Login page
    ├── signup.html          # Sign up with role selection
    ├── user_dashboard.html  # Donor dashboard (Donate / Need Blood)
    ├── donate.html          # Blood donation registration
    ├── need_blood.html      # Find blood by type in Chennai
    ├── wishlist.html        # User wishlist
    ├── hospital_dashboard.html # Hospital blood inventory
    ├── upload_blood.html    # Hospital upload form
    └── about.html           # Project info & module breakdown
```

## Setup

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Run the app
```bash
python app.py
```

### 3. Open browser
```
http://localhost:5000
```

## Excel File (blood_data.xlsx)
Automatically created on first run with 5 sheets:
- **Users** — Registered users (donors and hospitals)
- **Donations** — All blood donation registrations
- **BloodRequests** — Blood need requests
- **Wishlist** — User blood type wishlists
- **HospitalBlood** — Hospital blood inventory uploads

## Key Rules
- Users can only donate once every **8 weeks (56 days)**
- Hospital accounts can upload blood availability
- Regular users can donate or search for blood
- Wishlist items can be added when searching for blood by type

## Chennai Blood Banks Included
1. Rotary TTK Blood Bank - Adyar
2. Voluntary Health Services Blood Bank - Adyar
3. Government General Hospital - Park Town
4. Apollo Hospital Blood Bank - Greams Road
5. Fortis Malar Hospital - Adyar
6. Sri Ramachandra Medical Centre - Porur
7. Madras Medical Mission - Mogappair
8. Stanley Medical College - Royapuram
9. Kilpauk Medical College - Kilpauk
10. JIPMER Chennai - Anna Nagar
