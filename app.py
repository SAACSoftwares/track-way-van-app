import random
import json
import calendar
from flask import Flask, request, jsonify, render_template_string
from flask_cors import CORS
from datetime import datetime
from pathlib import Path
import json, openpyxl

app = Flask(__name__, static_folder='static', template_folder='templates')
CORS(app)

from flask import render_template

@app.route('/')
def serve_index():
    return render_template('index.html')
from flask import render_template

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/trip_logger")
def trip_logger():
    return render_template("trip_logger.html")

@app.route("/view_logs")
def view_logs_page():
    return render_template("view_logs.html")

@app.route("/attendance")
def attendance():
    return render_template("attendance.html")

@app.route("/attendance_viewer")
def attendance_viewer():
    return render_template("attendance_viewer.html")

@app.route("/admin_login")
def admin_login_page():
    return render_template("admin_login.html")

@app.route("/admin_dashboard")
def admin_dashboard():
    return render_template("admin_dashboard.html")

@app.route("/change_password")
def change_password_page():
    return render_template("change_password.html")

@app.route("/forgot_password")
def forgot_password():
    return render_template("forgot_password.html")

@app.route("/admin_email")
def admin_email():
    return render_template("admin_email.html")



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

# === Utility: Trip Logger Excel ===
def get_excel_path():
    filename = datetime.now().strftime("%B_%Y").lower() + ".xlsx"
    path = Path(filename)
    if not path.exists():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Trips"
        ws.append([
            "Van ID", "Driver Name", "Route", "Contact", "Seat Capacity",
            "Departure Time", "Arrival Time", "Odometer Start", "Odometer End", "Distance"
        ])
        wb.save(path)
    return path

# === API: Fetch Van Details ===
@app.route("/get_van_details", methods=["GET"])
def get_van_details():
    van_id = request.args.get("van_id", "").strip().upper()
    details = van_data.get(van_id)
    if details:
        return jsonify({
            "van_id": van_id,
            "route": details["route"],
            "driver": details["driver"],
            "contact": details["contact"],
            "seat_capacity": details["seat_capacity"],
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })
    return jsonify({"error": "Van not found"}), 404

# === API: Log Trip ===
@app.route("/log_trip", methods=["POST"])
def log_trip():
    data = request.get_json()
    van_id = data.get("van_id", "").strip().upper()
    stage = data.get("stage", "").strip().lower()
    odometer = data.get("odometer")

    if not van_id or stage not in ["departure", "arrival"] or odometer is None:
        return jsonify({"error": "Invalid data"}), 400

    try:
        odo = int(odometer)
    except:
        return jsonify({"error": "Odometer must be a number"}), 400

    details = van_data.get(van_id)
    if not details:
        return jsonify({"error": "Van not found in data"}), 404

    excel_path = get_excel_path()
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if stage == "departure":
        ws.append([
            van_id, details["driver"], details["route"], details["contact"],
            details["seat_capacity"], now, "", odo, "", ""
        ])
        wb.save(excel_path)
        send_file_to_admin(excel_path, "Trip Log Update", f"Updated trip log for van {van_id}.")
        return jsonify({"message": "✅ Departure logged successfully"})

    elif stage == "arrival":
        for row in reversed(list(ws.iter_rows(min_row=2))):
            if str(row[0].value).strip().upper() == van_id and row[6].value in (None, ""):
                row[6].value = now
                row[8].value = odo
                try:
                    distance = int(odo) - int(row[7].value)
                    row[9].value = distance
                except:
                    row[9].value = ""
                wb.save(excel_path)
                send_file_to_admin(excel_path, "Trip Log Update", f"Updated trip log for van {van_id}.")
                return jsonify({"message": "✅ Arrival logged successfully"})
        return jsonify({"error": "No matching departure found"}), 400

# === API: Get Students for a Van ===
@app.route("/get_students", methods=["GET"])
def get_students():
    van_id = request.args.get("van_id", "").strip().upper()
    students = student_data.get(van_id)
    if not students:
        return jsonify({"error": "No students found for this van."}), 404
    return jsonify(students)

# === API: Submit Attendance ===
@app.route("/submit_attendance", methods=["POST"])
def submit_attendance():
    data = request.get_json()
    van_id = data.get("van_id", "").strip().upper()
    records = data.get("records")

    if not van_id or not records:
        return jsonify({"error": "Invalid data"}), 400

    today_str = datetime.now().strftime("%d-%b")
    filename = f"attendance_{van_id.lower()}_{datetime.now().strftime('%B_%Y').lower()}.xlsx"
    path = Path(filename)

    if not path.exists():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Attendance"
        headers = ["S.No", "Name", "Class & Section", "Address", "Parent Contact"]
        ws.append(headers + [today_str])
        for student in student_data.get(van_id, []):
            ws.append([
                student["sno"], student["name"], student["class"],
                student["address"], student["parent_contact"], ""
            ])
    else:
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        headers = [cell.value for cell in ws[1]]
        if today_str not in headers:
            ws.cell(row=1, column=len(headers) + 1).value = today_str

    date_col = [cell.value for cell in ws[1]].index(today_str) + 1

    for record in records:
        sno = record.get("sno")
        status = record.get("status")
        for row in ws.iter_rows(min_row=2):
            if row[0].value == sno:
                row[date_col - 1].value = status
                break

    wb.save(path)
    send_file_to_admin(path, "Attendance Sheet Update", f"Updated attendance for van {van_id}.")
    return jsonify({"message": "✅ Attendance saved!"})

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

# Get security question
@app.route("/get_security_question")
def get_security_question():
    try:
        with open(CONFIG_FILE, "r") as f:
            config = json.load(f)
        return jsonify({"question": config.get("security_question", "No question set.")})
    except:
        return jsonify({"question": "Unable to load question."})

# Forgot Password - verify answer and reset password
@app.route("/forgot_password", methods=["POST"])
def forgot_password():
    data = request.get_json()
    answer = data.get("answer", "").strip().lower()
    new_password = data.get("new_password", "").strip()

    try:
        with open(CONFIG_FILE, "r") as f:
            config = json.load(f)

        if answer != config.get("security_answer", "").strip().lower():
            return jsonify({"success": False, "error": "Incorrect answer."}), 401

        config["password"] = new_password
        with open(CONFIG_FILE, "w") as f:
            json.dump(config, f, indent=2)

        return jsonify({"success": True, "message": "✅ Password reset successfully."})
    except:
        return jsonify({"success": False, "error": "Error updating password."}), 500
@app.route("/update_admin_email", methods=["POST"])
def update_admin_email():
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
    return jsonify(van_data)

@app.route("/update_van", methods=["POST"])
def update_van():
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
import smtplib
import os
from email.mime.text import MIMEText
from flask import session
from dotenv import load_dotenv

load_dotenv()

OTP_STORE = {}  # {email: otp}
ADMIN_EMAIL_OTP_STORE = {}
ADMIN_EMAIL = os.getenv("ADMIN_EMAIL")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD")  # for SMTP login
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
        with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
            smtp.starttls()
            smtp.login(ADMIN_EMAIL, ADMIN_PASSWORD)
            smtp.send_message(msg)
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

    # Prepare email
    msg = MIMEText(f"Your OTP to reset password is: {otp}")
    msg["Subject"] = "Track Way - Password Reset OTP"
    msg["From"] = ADMIN_EMAIL
    msg["To"] = email

    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
            smtp.starttls()
            smtp.login(ADMIN_EMAIL, ADMIN_PASSWORD)
            smtp.send_message(msg)
        return jsonify({"success": True, "message": "✅ OTP sent to your email."})
    except Exception as e:
        return jsonify({"error": f"Failed to send email: {str(e)}"}), 500

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
    

@app.route("/send_admin_email_otp", methods=["POST"])
def send_admin_email_otp():
    data = request.get_json()
    email = data.get("email", "").strip()
    if "@" not in email:
        return jsonify({"success": False, "error": "Invalid email"}), 400

    otp = str(random.randint(100000, 999999))
    ADMIN_EMAIL_OTP_STORE[email] = otp

    msg = MIMEText(f"Your OTP to confirm admin email is: {otp}")
    msg["Subject"] = "Track Way - Confirm Admin Email"
    msg["From"] = ADMIN_EMAIL
    msg["To"] = email

    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
            smtp.starttls()
            smtp.login(ADMIN_EMAIL, ADMIN_PASSWORD)
            smtp.send_message(msg)
        return jsonify({"success": True, "message": "✅ OTP sent to new email."})
    except Exception as e:
        return jsonify({"success": False, "error": f"Failed to send OTP: {str(e)}"}), 500

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


@app.route("/healthz")
def health_check():
    return "OK", 200


import os

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, host="0.0.0.0", port=port)
