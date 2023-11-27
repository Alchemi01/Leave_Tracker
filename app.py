from flask import Flask, render_template, request, redirect, url_for
from flask_mail import Mail, Message
import openpyxl
from datetime import datetime

app = Flask(__name__)

# Configure Flask-Mail
app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USERNAME'] = 'okeyodemichael@gmail.com'
app.config['MAIL_PASSWORD'] = 'rdzk deus nbgh tiwu'
app.config['MAIL_DEFAULT_SENDER'] = 'okeyodemichael@gmail.com'

mail = Mail(app)

@app.route('/')
def index():
    return render_template('leave_application_form.html')

@app.route('/submit_leave', methods=['POST'])
def submit_leave():
    employee_id = request.form['employee_id']
    start_date = request.form['start_date']
    end_date = request.form['end_date']
    supervisor_email = request.form['supervisor_email']

    # Send email to supervisor
    send_email(employee_id, start_date, end_date, supervisor_email)

    # Save record to Excel
    #save_to_excel(employee_id, start_date, end_date)

    return redirect(url_for('index'))

def send_email(employee_id, start_date, end_date, supervisor_email):
    subject = f"Leave Application - Employee {employee_id}"
    body = f"Leave requested for Employee {employee_id} from {start_date} to {end_date}. Please review and approve."

    msg = Message(subject, recipients=[supervisor_email], body=body)
    mail.send(msg)



if __name__ == '__main__':
    app.run(debug=True)
