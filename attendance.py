import cv2
import numpy as np
import face_recognition
import pandas as pd
import datetime
import os
import time
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Student emails
student_emails = {
    "karthik": "karthikvemula62@gmail.com",
    "santosh": "22kq1a0757@pace.ac.in",
    "student3": "student3@example.com",
    "arun": "22kq1a0753@pace.ac.in"
}

# Load known student images
known_faces_dir = "students/"
unknown_faces_dir = "unknown_faces/"
os.makedirs(unknown_faces_dir, exist_ok=True)

known_encodings = []
student_names = []

for filename in os.listdir(known_faces_dir):
    if filename.endswith(".jpg") or filename.endswith(".png"):
        image = face_recognition.load_image_file(os.path.join(known_faces_dir, filename))
        encoding = face_recognition.face_encodings(image)[0]
        known_encodings.append(encoding)
        student_names.append(os.path.splitext(filename)[0])

# Initialize attendance file
today_date = datetime.datetime.now().strftime("%Y-%m-%d")
file_path = f"attendance_{today_date}.xlsx"
try:
    df = pd.read_excel(file_path)
except FileNotFoundError:
    df = pd.DataFrame(columns=["Name", "Time", "Date"])
    df.to_excel(file_path, index=False)

# Start webcam
video = cv2.VideoCapture(0)
start_time = time.time()

while True:
    ret, frame = video.read()
    rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

    face_locations = face_recognition.face_locations(rgb_frame)
    face_encodings = face_recognition.face_encodings(rgb_frame, face_locations)

    for face_encoding, location in zip(face_encodings, face_locations):
        name = "Unknown"
        face_distances = face_recognition.face_distance(known_encodings, face_encoding)

        best_match_index = np.argmin(face_distances)
        if face_distances[best_match_index] < 0.6:
            name = student_names[best_match_index]

            now = datetime.datetime.now()
            date_str = now.strftime("%Y-%m-%d")
            time_str = now.strftime("%H:%M:%S")

            if not ((df["Name"] == name) & (df["Date"] == date_str)).any():
                df = pd.concat([df, pd.DataFrame([[name, time_str, date_str]], columns=df.columns)], ignore_index=True)
                df.to_excel(file_path, index=False)
        else:
            top, right, bottom, left = location
            face_image = frame[top:bottom, left:right]
            unknown_filename = os.path.join(unknown_faces_dir, f"unknown_{datetime.datetime.now().strftime('%H%M%S')}.jpg")
            cv2.imwrite(unknown_filename, face_image)

        top, right, bottom, left = location
        color = (0, 255, 0) if name != "Unknown" else (0, 0, 255)
        cv2.rectangle(frame, (left, top), (right, bottom), color, 2)
        cv2.putText(frame, name, (left, top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.8, color, 2)

    cv2.imshow("Smart Attendance System", frame)

    if time.time() - start_time > 20:
        print("Auto shutting down the camera after 30 seconds.")
        break

    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

video.release()
cv2.destroyAllWindows()

# Show today's attendance
print("\n--- Today's Attendance ---")
today_records = df[df["Date"] == today_date]
print(today_records[["Name", "Time"]])

# Attendance percentage
all_data = pd.concat([pd.read_excel(f) for f in os.listdir() if f.startswith("attendance_") and f.endswith(".xlsx")])
total_days = len(all_data["Date"].unique())
student_attendance = all_data.groupby("Name")["Date"].count()
attendance_percentage = (student_attendance / total_days) * 100

attendance_percentage_df = attendance_percentage.reset_index()
attendance_percentage_df.columns = ["Name", "Attendance Percentage"]
attendance_percentage_df.to_excel("attendance_percentage.xlsx", index=False)
print("\nAttendance percentage saved to attendance_percentage.xlsx")

# Email for absent students
def send_email(to_email, student_name):
    subject = "Attendance Notification"
    body = f"Hi {student_name},\n\nYou student were marked absent today in class.\n\nRegards,\nSmart Attendance System"
    
    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = "karthikvemula62@gmail.com"
    msg["To"] = to_email

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login("karthikvemula62@gmail.com", "thev lspo rlpf lsyx")
        server.send_message(msg)

present_today = today_records["Name"].tolist()
absent_students = [name for name in student_emails.keys() if name not in present_today]

for student in absent_students:
    send_email(student_emails[student], student)
    print(f"Sent email to {student_emails[student]}")

# Send full Excel sheet to admin
def send_admin_excel(admin_email):
    subject = f"Attendance Report - {today_date}"
    body = "Hi Admin,\n\nPlease find attached the attendance report for today.\n\nRegards,\nSmart Attendance System"

    msg = MIMEMultipart()
    msg["From"] = "karthikvemula62@gmail.com"
    msg["To"] = admin_email
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain"))

    with open(file_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename= {os.path.basename(file_path)}")
        msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login("karthikvemula62@gmail.com", "thev lspo rlpf lsyx")
        server.send_message(msg)

send_admin_excel("karthikvemula62@gmail.com")
print("✅ Attendance report sent to admin.")
