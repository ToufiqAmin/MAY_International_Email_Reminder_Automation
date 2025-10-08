# 📬 Email Reminder Automation

This project automates email reminders for upcoming events using Python. It sends two reminders:
- ✅ 7 days before the event
- ✅ 2 days before the event

It uses a CSV file to track events and securely sends emails using Gmail SMTP.

---

## 📁 Folder Structure
email_reminder/
├── .env                    # Secure credentials
├── events.xlsx              # Event schedule
├── reminder.py             # Main automation script
├── requirements.txt        # Dependency lock file
├── reminder.log            # Auto-generated log file
├── stop.flag               # Optional kill switch


---

## 🧪 Setup Instructions

### 1. Create Virtual Environment

```bash
python -m venv venv
venv\Scripts\activate  # Windows
source venv/bin/activate  # macOS/Linux

### 2. Install Dependencies
pip install -r requirements.txt

### 3. Create and Configure .env
EMAIL_ADDRESS=your_email@example.com
EMAIL_PASSWORD=your_app_password # Use Gmail App Passwords if 2FA is enabled.

### 4. Sample Excel
event_name,event_date,email
LC Expiry,2025-10-14,recipient1@example.com
Invoice Submission,2025-10-20,recipient2@example.com

### 5. Run The Automation
python reminder.py

### 6. Stop the Automation
echo. > stop.flag  # Windows CMD
touch stop.flag    # Git Bash

### 7. Restart the Automation
del stop.flag      # Windows
rm stop.flag       # macOS/Linux

## Features
- Secure credential storage via .env
- Logging of sent emails and errors
- Email format and date validation
- Kill switch for graceful shutdown
- Easy to extend with WhatsApp or SMS

## To Do
- Add GUI for event management
- Integrate WhatsApp reminders via Twilio
- Dashboard with pandas and matplotlib

## License
MIT License — free to use, modify, and distribute.
