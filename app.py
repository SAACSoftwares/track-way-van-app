import requests
import os
import random
import json
import calendar
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()
from flask import Flask, request, jsonify, render_template_string
from flask_cors import CORS
from datetime import datetime
from pathlib import Path
import json, openpyxl
from flask import session

from flask import redirect
import smtplib
from email.mime.text import MIMEText
def send_email_resend(to_email, subject, html):
    RESEND_API_KEY = os.getenv("RESEND_API_KEY")

    if not RESEND_API_KEY:
        raise Exception("RESEND_API_KEY not found")

    response = requests.post(
        "https://api.resend.com/emails",
        headers={
            "Authorization": f"Bearer {RESEND_API_KEY}",
            "Content-Type": "application/json"
        },
        json={
            "from": "Track Way <onboarding@resend.dev>",
            "to": [to_email],
            "subject": subject,
            "html": html
        },
        timeout=20
    )

    if response.status_code not in (200, 201):
        raise Exception(response.text)


def admin_required():
    return session.get("admin_logged_in") is True


app = Flask(__name__, static_folder='static', template_folder='templates')
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key")
CORS(app)

from flask import render_template

@app.route("/")
def home():
    return render_template("index.html")

@app.route('/trip_logger', methods=['GET', 'POST'])
def trip_logger():
    import json
    from datetime import datetime
    import os
    from openpyxl import load_workbook, Workbook
    import math

    def haversine(lat1, lon1, lat2, lon2):
        R = 6371000
        phi1, phi2 = math.radians(lat1), math.radians(lat2)
        d_phi = math.radians(lat2 - lat1)
        d_lambda = math.radians(lon2 - lon1)
        a = math.sin(d_phi/2)**2 + math.cos(phi1)*math.cos(phi2)*math.sin(d_lambda/2)**2
        return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))

    GATE_LAT = 17.312098  
    GATE_LNG = 76.814419 
    ALLOWED_RADIUS_METERS = 300

    with open("vans.json", "r") as f:
        vans = json.load(f)

    if request.method == 'POST':
        van_id = request.form.get('van_id')
        odometer = request.form.get('odometer')
        latitude = float(request.form.get('latitude', 0))
        longitude = float(request.form.get('longitude', 0))

        if van_id not in vans:
            return f"Van ID {van_id} not found.", 404

        distance = haversine(latitude, longitude, GATE_LAT, GATE_LNG)
        if distance > ALLOWED_RADIUS_METERS:
            return "You are not at the authorized logging location.", 403

        van_details = vans[van_id]
        now = datetime.now()
        month = now.strftime("%B").lower()
        year = now.year
        filename = f"{month}_{year}.xlsx"

        if os.path.exists(filename):
            wb = load_workbook(filename)
        else:
            wb = Workbook()
            ws = wb.active
            ws.append(["Van ID", "Van Number", "Arrival Odometer", "Arrival Time"])
        ws = wb.active
        ws.append([van_id, van_details['number'], odometer, now.strftime("%Y-%m-%d %H:%M:%S")])
        wb.save(filename)

        return f"Trip for {van_id} logged successfully."

    return render_template("trip_logger.html")


@app.route("/view_logs_page")
def view_logs_page():
    return render_template("view_logs.html")

@app.route("/attendance")
def attendance():
    return render_template("attendance.html")

@app.route("/attendance_viewer")
def attendance_viewer():
    return render_template("attendance_viewer.html")

@app.route("/admin_login_page")
def admin_login_page():
    return render_template("admin_login.html")


@app.route("/admin_dashboard")
def admin_dashboard():
    if not admin_required():
        return redirect("/admin_login")
    return render_template("admin_dashboard.html")

@app.route("/change_password")
def change_password_page():
    if not admin_required():
        return redirect("/admin_login")
    return render_template("change_password.html")

@app.route("/forgot_password")
def forgot_password_page():
    return render_template("forgot_password.html")





# Load van and student data
VAN_DATA_FILE = "vans.json"
STUDENT_DATA_FILE = "students.json"

try:
    with open(VAN_DATA_FILE, "r") as f:
        van_data = json.load(f)
except:
    van_data = {}

try:
    with open(STUDENT_DATA_FILE, "r") as f:
        student_data = json.load(f)
except:
    student_data = {}


# === API: Get Students for a Van ===
@app.route("/get_students", methods=["GET"])
def get_students():
    if not admin_required():
        return redirect("/admin_login")
    van_id = request.args.get("van_id", "").strip().upper()
    students = student_data.get(van_id)
    if not students:
        return jsonify({"error": "No students found for this van."}), 404
    return jsonify(students)

@app.route("/submit_attendance", methods=["POST"])
def submit_attendance():
    data = request.get_json()
    van_id = data.get("van_id", "").strip().upper()
    records = data.get("records")

    if not van_id or not records:
        return jsonify({"error": "Invalid data"}), 400

    now = datetime.now()
    today_str = now.strftime("%d-%b")
    month = now.strftime("%B").lower()
    year = now.strftime("%Y")

    filename = f"attendance_{van_id.lower()}_{month}_{year}.xlsx"
    path = Path(filename)

    # ---------------- CREATE FILE IF NOT EXISTS ----------------
    if not path.exists():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Attendance"

        headers = ["S.No", "Name", "Class & Section", "Address", "Parent Contact"]
        ws.append(headers + [today_str])

        for student in student_data.get(van_id, []):
            ws.append([
                student["sno"],
                student["name"],
                student["class"],
                student["address"],
                student["parent_contact"],
                ""
            ])

    # ---------------- UPDATE FILE IF EXISTS ----------------
    else:
        wb = openpyxl.load_workbook(path)
        ws = wb.active

        headers = [cell.value for cell in ws[1]]
        if today_str in headers:
            return jsonify({"error": "Attendance already marked for today"}), 400


        if today_str not in headers:
            ws.cell(row=1, column=len(headers) + 1).value = today_str
            headers.append(today_str)

    # ---------------- FIND DATE COLUMN ----------------
    headers = [cell.value for cell in ws[1]]
    date_col = headers.index(today_str) + 1

    # ---------------- WRITE ATTENDANCE ----------------
    for record in records:
        sno = record.get("sno")
        status = record.get("status")

        for row in ws.iter_rows(min_row=2):
            if row[0].value == sno:
                row[date_col - 1].value = status
                break

    wb.save(path)
    return jsonify({"message": "✅ Attendance saved successfully!"})


@app.route("/send_attendance_summary", methods=["POST"])
def send_attendance_summary():
    if not admin_required():
        return redirect("/admin_login")
    data = request.get_json()
    van_id = data.get("van_id", "").strip().lower()
    month = data.get("month", "").strip().lower()
    year = data.get("year", "").strip()

    if not van_id or not month or not year:
        return jsonify({"error": "Missing van_id, month, or year"}), 400

    filename = f"attendance_{van_id.lower()}_{month}_{year}.xlsx"
    path = Path(filename)

    if path.exists():
        send_file_to_admin(path, f"📋 Attendance Summary for {van_id.upper()}", f"Attached is the attendance summary for {van_id.upper()} - {month.capitalize()} {year}.")
        return jsonify({"message": "✅ Attendance summary sent via email."})
    return jsonify({"error": "Attendance file not found"}), 404

# === View Trip Logs in Table Format ===
view_logs_template = """
<!DOCTYPE html>
<html lang='en'>
<head>
  <meta charset='UTF-8'>
  <title>Van Trip Logs - {{ month }} {{ year }}</title>
  <style>
    body { font-family: Arial; background: #f0f0f0; padding: 20px; }
    table { width: 100%; border-collapse: collapse; background: white; }
    th, td { padding: 10px; border: 1px solid #ccc; text-align: center; }
    th { background: #333; color: white; }
    h2 { text-align: center; }
    .footer { margin-top: 40px; text-align: center; font-size: 14px; color: #555; }
  </style>
</head>
<body>
  <h2>Van Trip Logs - {{ month|capitalize }} {{ year }}</h2>
  {% if rows %}
  <table>
    <thead>
      <tr>{% for h in headers %}<th>{{ h }}</th>{% endfor %}</tr>
    </thead>
    <tbody>
      {% for row in rows %}<tr>{% for cell in row %}<td>{{ cell }}</td>{% endfor %}</tr>{% endfor %}
    </tbody>
  </table>
  {% else %}
  <p style="text-align: center;">No logs found for this month.</p>
  {% endif %}
  <div class="footer">
    Contact us: <a href="mailto:adityashetty2908@gmail.com">adityshetty2908@gmail.com</a>
  </div>
</body>
</html>
"""

@app.route("/view_logs")
def view_logs():
    month = request.args.get("month", "").lower()
    year = request.args.get("year", "").strip()
    if not month or not year:
        return "Provide month and year, e.g., /view_logs?month=july&year=2025", 400

    filename = f"{month}_{year}.xlsx"
    path = Path(filename)
    if not path.exists():
        return render_template_string(view_logs_template, headers=[], rows=[], month=month, year=year)

    wb = openpyxl.load_workbook(path)
    ws = wb.active
    data = list(ws.iter_rows(values_only=True))
    headers, rows = data[0], data[1:]
    return render_template_string(view_logs_template, headers=headers, rows=rows, month=month, year=year)

@app.route("/view_attendance", methods=["GET"])
def view_attendance():
    van_id = request.args.get("van_id", "").strip().upper()
    month = request.args.get("month", "").strip().lower()
    year = request.args.get("year", "").strip()

    if not van_id or not month or not year:
        return "Missing parameters", 400

    filename = f"attendance_{van_id.lower()}_{month}_{year}.xlsx"
    path = Path(filename)
    if not path.exists():
        return f"No attendance file found for {van_id} - {month.capitalize()} {year}", 404

    wb = openpyxl.load_workbook(path)
    ws = wb.active
    data = list(ws.iter_rows(values_only=True))
    headers, rows = data[0], data[1:]

    # Render as HTML table
    html = f"<h2 style='text-align:center;'>{van_id} - Attendance ({month.capitalize()} {year})</h2><table border='1' style='width:100%;border-collapse:collapse;text-align:center;'>"
    html += "<tr>" + "".join(f"<th>{h}</th>" for h in headers) + "</tr>"
    for row in rows:
        html += "<tr>" + "".join(f"<td>{cell if cell is not None else ''}</td>" for cell in row) + "</tr>"
    html += "</table>"

    return html
from werkzeug.security import check_password_hash

# Store the hashed password (this should ideally be in a secure config file)
ADMIN_USERNAME = "admin"
ADMIN_HASHED_PASSWORD = "scrypt:32768:8:1$UH3CtH4I4KdNfOtG$e9410b70406888c1ae27694b3a61e57e22a4ae7fbfc9a5dfe91f18a6021fae2a2d137f6f0f304fe9e63e86f9f10e01a95edf7c02bc127a166866c8a96511caea"

@app.route("/admin_login", methods=["POST"])
def admin_login():
    data = request.get_json()
    username = data.get("username")
    password = data.get("password")
    stored_hash = load_admin_password()

    if not stored_hash:
        return jsonify({"success": False, "message": "Password file missing or unreadable"}), 500

    if username == ADMIN_USERNAME and check_password_hash(stored_hash, password):
        session["admin_logged_in"] = True
        return jsonify({"success": True})

    else:
        return jsonify({"success": False, "message": "Invalid username or password"}), 401

from werkzeug.security import generate_password_hash

PASSWORD_FILE = "admin_password.json"

def load_admin_password():
    try:
        with open("admin_password.json", "r") as f:
            return json.load(f)["hashed"]
    except Exception as e:
        print("❌ Failed to load password file:", e)
        return None


def save_admin_password(new_hash):
    with open("admin_password.json", "w") as f:
        json.dump({"hashed": new_hash}, f)

@app.route("/change_password", methods=["POST"])
def change_password():
    if not admin_required():
        return redirect("/admin_login")
    data = request.get_json()
    old_pass = data.get("old_password")
    new_pass = data.get("new_password")

    stored_hash = load_admin_password()

    if not check_password_hash(stored_hash, old_pass):
        return jsonify({"error": "❌ Old password is incorrect"}), 400

    new_hash = generate_password_hash(new_pass, method="scrypt")
    save_admin_password(new_hash)
    return jsonify({"message": "✅ Password updated successfully"})
import json

CONFIG_FILE = "config.json"



@app.route("/update_admin_email", methods=["POST"])
def update_admin_email():
    if not admin_required():
        return redirect("/admin_login")
    data = request.get_json()
    new_email = data.get("email", "").strip()

    if not new_email or "@" not in new_email:
        return jsonify({"error": "Invalid email"}), 400

    try:
        with open("config.json", "r") as f:
            config = json.load(f)
    except:
        config = {}

    config["email"] = new_email

    with open("config.json", "w") as f:
        json.dump(config, f, indent=2)

    return jsonify({"message": "✅ Admin email updated successfully."})
@app.route("/get_admin_email", methods=["GET"])
def get_admin_email():
    try:
        with open("config.json", "r") as f:
            config = json.load(f)
        return jsonify({"email": config.get("email", "")})
    except:
        return jsonify({"email": ""})
# === Add Van Entry ===
@app.route("/add_van", methods=["POST"])
def add_van():
    if not admin_required():
        return redirect("/admin_login")
    data = request.get_json()
    van_id = data.get("van_id", "").strip().upper()

    if not van_id or van_id in van_data:
        return jsonify({"error": "Invalid or existing van ID"}), 400

    van_data[van_id] = {
        "route": data.get("route", ""),
        "driver": data.get("driver", ""),
        "contact": data.get("contact", ""),
        "seat_capacity": data.get("seat_capacity", "")
    }

    with open("vans.json", "w") as f:
        json.dump(van_data, f, indent=2)

    return jsonify({"message": f"✅ Van {van_id} added successfully!"})
# === Add Student to a Van ===
@app.route("/add_student", methods=["POST"])
def add_student():
    if not admin_required():
        return redirect("/admin_login")
    data = request.get_json()
    van_id = data.get("van_id", "").strip().upper()

    if not van_id:
        return jsonify({"error": "Van ID required"}), 400

    # Initialize if van doesn't exist in student list
    if van_id not in student_data:
        student_data[van_id] = []

    student_list = student_data[van_id]
    sno = len(student_list) + 1

    student_list.append({
        "sno": sno,
        "name": data.get("name", ""),
        "class": data.get("class", ""),
        "address": data.get("address", ""),
        "parent_contact": data.get("parent_contact", "")
    })

    with open("students.json", "w") as f:
        json.dump(student_data, f, indent=2)

    return jsonify({"message": f"✅ Student added to van {van_id}!"})

@app.route("/get_all_vans")
def get_all_vans():
    if not admin_required():
        return redirect("/admin_login")
    return jsonify(van_data)

@app.route("/update_van", methods=["POST"])
def update_van():
    if not admin_required():
        return redirect("/admin_login")
    data = request.get_json()
    van_id = data.get("van_id", "").strip().upper()

    if van_id not in van_data:
        return jsonify({"error": "Van not found"}), 404

    van_data[van_id] = {
        "route": data.get("route", ""),
        "driver": data.get("driver", ""),
        "contact": data.get("contact", ""),
        "seat_capacity": data.get("seat_capacity", "")
    }

    with open("vans.json", "w") as f:
        json.dump(van_data, f, indent=2)

    return jsonify({"message": f"✅ Van {van_id} updated successfully!"})

@app.route("/delete_van", methods=["POST"])
def delete_van():
    if not admin_required():
        return redirect("/admin_login")
    data = request.get_json()
    van_id = data.get("van_id", "").strip().upper()

    if van_id in van_data:
        van_data.pop(van_id)
        with open("vans.json", "w") as f:
            json.dump(van_data, f, indent=2)
        return jsonify({"message": f"🗑️ Van {van_id} deleted successfully!"})

    return jsonify({"error": "Van not found"}), 404
@app.route("/update_student", methods=["POST"])
def update_student():
    if not admin_required():
        return redirect("/admin_login")
    data = request.get_json()
    van_id = data.get("van_id", "").strip().upper()
    sno = data.get("sno")

    if van_id not in student_data or not sno:
        return jsonify({"error": "Invalid van or student"}), 400

    for student in student_data[van_id]:
        if student["sno"] == sno:
            student["name"] = data.get("name", "")
            student["class"] = data.get("class", "")
            student["address"] = data.get("address", "")
            student["parent_contact"] = data.get("parent_contact", "")
            break

    with open("students.json", "w") as f:
        json.dump(student_data, f, indent=2)

    return jsonify({"message": f"✅ Student S.No {sno} updated in {van_id}"})

@app.route("/delete_student", methods=["POST"])
def delete_student():
    if not admin_required():
        return redirect("/admin_login")
    data = request.get_json()
    van_id = data.get("van_id", "").strip().upper()
    sno = data.get("sno")

    if van_id not in student_data or not sno:
        return jsonify({"error": "Invalid request"}), 400

    original = len(student_data[van_id])
    student_data[van_id] = [s for s in student_data[van_id] if s["sno"] != sno]

    # Reassign serial numbers
    for idx, student in enumerate(student_data[van_id], 1):
        student["sno"] = idx

    with open("students.json", "w") as f:
        json.dump(student_data, f, indent=2)

    return jsonify({"message": f"🗑️ Student S.No {sno} removed from {van_id}"})

import os

from flask import session

OTP_STORE = {}  # {email: otp}
ADMIN_EMAIL_OTP_STORE = {}
ADMIN_EMAIL = os.getenv("ADMIN_EMAIL")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD")  # for SMTP login

print("🔐 ADMIN_EMAIL =", ADMIN_EMAIL)
print("🔐 ADMIN_PASSWORD loaded =", bool(ADMIN_PASSWORD))

from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

def send_file_to_admin(filepath, subject, body):
    if not os.path.exists(filepath):
        print("❌ File not found:", filepath)
        return

    try:
        with open("config.json", "r") as f:
            config = json.load(f)
        to_email = config.get("email", ADMIN_EMAIL)
    except:
        to_email = ADMIN_EMAIL

    msg = MIMEMultipart()
    msg["From"] = ADMIN_EMAIL
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    with open(filepath, "rb") as file:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(filepath)}"')
        msg.attach(part)

    try:
        print(f"📤 Sent {os.path.basename(filepath)} to {to_email}")
    except Exception as e:
        print(f"❌ Email send failed: {str(e)}")

@app.route("/send_otp", methods=["POST"])
def send_otp():
    data = request.get_json()
    email = data.get("email")

    if not email:
        return jsonify({"error": "Email required"}), 400

    otp = str(random.randint(100000, 999999))
    OTP_STORE[email] = otp

    try:
        send_email_resend(
            to_email=email,
            subject="Track Way - Password Reset OTP",
            html=f"""
                <h2>Password Reset OTP</h2>
                <p>Your OTP is:</p>
                <h1>{otp}</h1>
                <p>This OTP is valid for a short time.</p>
            """
        )

        return jsonify({"success": True, "message": "OTP sent successfully"})

    except Exception as e:
        print("❌ Resend OTP failed:", e)
        return jsonify({"error": "Failed to send OTP"}), 500



@app.route("/reset_password", methods=["POST"])
def reset_password():
    data = request.get_json()
    email = data.get("email")
    otp = data.get("otp")
    new_password = data.get("new_password")

    if not email or not otp or not new_password:
        return jsonify({"error": "Missing fields"}), 400

    if OTP_STORE.get(email) == otp:
        new_hash = generate_password_hash(new_password, method="scrypt")
        save_admin_password(new_hash)
        OTP_STORE.pop(email)
        return jsonify({"success": True, "message": "✅ Password updated successfully."})
    else:
        return jsonify({"error": "Invalid OTP or email"}), 400
    



@app.route("/verify_admin_email_otp", methods=["POST"])
def verify_admin_email_otp():
    data = request.get_json()
    email = data.get("email", "").strip()
    otp = data.get("otp", "").strip()

    if ADMIN_EMAIL_OTP_STORE.get(email) == otp:
        try:
            with open("config.json", "r") as f:
                config = json.load(f)
        except:
            config = {}

        config["email"] = email
        with open("config.json", "w") as f:
            json.dump(config, f, indent=2)

        ADMIN_EMAIL_OTP_STORE.pop(email, None)
        return jsonify({"success": True, "message": "✅ Admin email updated!"})
    else:
        return jsonify({"success": False, "error": "Incorrect OTP"}), 400

from flask import jsonify
from datetime import datetime

@app.route("/send_daily_trip_log", methods=["POST"])
def send_daily_trip_log():
    if not admin_required():
        return redirect("/admin_login")
    today = datetime.now().strftime("%Y-%m-%d")
    filename = datetime.now().strftime("%B_%Y").lower() + ".xlsx"
    path = Path(filename)

    if not path.exists():
        return jsonify({"error": "Trip log file not found"}), 404

    send_file_to_admin(path, f"🚌 Daily Trip Log - {today}", f"Attached is the trip log for {today}.")
    return jsonify({"message": "✅ Daily trip log sent successfully."})

@app.route("/healthz")
def health_check():
    return "OK", 200

# Add this import at the top if not already present
from flask import render_template

# Add this route near your existing routes
@app.route('/scan')
def scan_qr():
    return render_template('scan.html')

import os

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, host="0.0.0.0", port=port)
