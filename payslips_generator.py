
import pandas as pd
from fpdf import FPDF
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import ssl
import os
import random
from datetime import datetime

def generate_random_phone():
    return f"+44 7{random.randint(100,999)} {random.randint(100,999)} {random.randint(1000,9999)}"

def main():
    try:
        payslips_dir = r'C:\Users\uncommonStudent\Desktop\Payslip\payslips'
        os.makedirs(payslips_dir, exist_ok=True)
        if not os.access(payslips_dir, os.W_OK):
            print("Directory is not writable – check permissions")
            return

        df = pd.read_excel('employees.xlsx.xlsx')

        # Clean and calculate Net Salary
        df['BASIC SALARY'] = df['BASIC SALARY'].replace(',', '', regex=True).astype(float)
        df['ALLOWANCES'] = df['ALLOWANCES'].replace(',', '', regex=True).astype(float)
        df['DEDUCTIONS'] = df['DEDUCTIONS'].replace(',', '', regex=True).astype(float)
        df['Net Salary'] = df['BASIC SALARY'] + df['ALLOWANCES'] - df['DEDUCTIONS']

        def generate_payslip(employee):
            try:
                print(f"Generating payslip for {employee['NAME']}...")
                pdf = FPDF(format='A4')
                pdf.add_page()
                try:
                    pdf.image('company_logo.png', x=50, y=10, w=100)
                except:
                    print("Company logo not found")
                pdf.set_line_width(0.5)
                pdf.set_fill_color(255, 165, 0)
                pdf.rect(10, 60, 190, 10, style='F')
                pdf.set_text_color(255, 255, 255)
                pdf.set_font('Arial', 'B', 16)
                pdf.cell(200, 10, 'AlvisTacool Automotive', ln=True, align='C')
                pdf.set_text_color(0, 0, 0)
                pdf.set_font('Arial', 'I', 12)
                pdf.cell(200, 10, 'Payroll Department', ln=True, align='C')
                pdf.ln(15)

                pdf.set_font('Arial', '', 10)
                company_info = [
                    'Address: 123 Main Street, London, UK',
                    'Phone: ' + generate_random_phone(),
                    'Email: payroll@alvistacool.com'
                ]
                for info in company_info:
                    pdf.cell(200, 10, info, ln=True, align='L')
                pdf.ln(10)

                pdf.set_font('Arial', 'B', 14)
                pdf.cell(200, 10, 'Employee Payslip', ln=True, align='C')
                pdf.set_font('Arial', '', 12)
                details = [
                    ('Employee Name:', employee['NAME']),
                    ('Employee Number:', employee['EMPLOYEE ID']),
                    ('Pay Period:', datetime.now().strftime('%d-%B-%Y'))
                ]
                x, y = 15, 100
                for key, value in details:
                    pdf.set_xy(x, y)
                    pdf.cell(100, 10, f"{key} {value}", align='L')
                    y += 10
                    if y > 250:
                        y = pdf.get_y()
                        x = 120

                pdf.ln(20)
                pdf.set_font('Arial', 'B', 14)
                pdf.cell(200, 10, 'Salary Breakdown:', ln=True, align='L')
                pdf.set_line_width(0.5)
                pdf.set_draw_color(255, 165, 0)
                pdf.line(10, pdf.get_y(), 200, pdf.get_y())

                salary_details = [
                    ('Basic Salary:', f"{employee['BASIC SALARY']:,.2f}"),
                    ('Allowances:', f"{employee['ALLOWANCES']:,.2f}"),
                    ('Total Deductions:', f"{employee['DEDUCTIONS']:,.2f}"),
                    ('Net Salary:', f"{employee['Net Salary']:,.2f}")
                ]
                x, y = 15, pdf.get_y() + 10
                for key, value in salary_details:
                    pdf.set_xy(x, y)
                    pdf.cell(100, 10, f"{key} {value}", align='L')
                    y += 10
                    if y > 250:
                        y = pdf.get_y()
                        x = 120

                pdf.ln(20)
                pdf.set_line_width(0.5)
                pdf.set_fill_color(255, 165, 0)
                pdf.rect(10, 260, 190, 10, style='F')
                pdf.set_text_color(255, 255, 255)
                pdf.set_font('Arial', 'I', 10)
                pdf.cell(200, 10, 'This payslip is computer generated and does not require a signature.', ln=True, align='C')

                # Save to target directory
                filename = f"Payslip_{employee['EMPLOYEE ID']}.pdf"
                pdf_path = os.path.join(payslips_dir, filename)
                pdf.output(pdf_path)
                print(f"Saved PDF to: {pdf_path}")
            except Exception as e:
                print(f"Error generating payslip: {e}")

        for _, row in df.iterrows():
            generate_payslip(row)

        # Email sending section
        def send_payslip_email(sender_email, app_password, recipient_email, payslip_path):
            try:
                context = ssl.create_default_context()
                with smtplib.SMTP('smtp.gmail.com', 587) as server:
                    server.starttls(context=context)
                    server.login(sender_email, app_password)
                    message = MIMEMultipart()
                    message['From'] = sender_email
                    message['To'] = recipient_email
                    message['Subject'] = 'Your Payslip for This Month'

                    body = f"""
                    Dear {recipient_email.split('@')[0].capitalize()},
                    Please find your payslip attached to this email.
                    Best regards,
                    Payroll Department
                    """
                    message.attach(MIMEText(body, 'plain'))

                    with open(payslip_path, 'rb') as f:
                        attach = MIMEApplication(f.read(), _subtype='pdf')
                        attach.add_header('Content-Disposition', 'attachment', filename=os.path.basename(payslip_path))
                        message.attach(attach)

                    server.send_message(message)
                    print(f"Email sent to {recipient_email}")
            except Exception as e:
                print(f"Email error for {recipient_email}: {e}")

        sender_email = ""
        app_password = ""  # Make sure this is stored securely!
        for _, row in df.iterrows():
            payslip_path = os.path.join(payslips_dir, f"Payslip_{row['EMPLOYEE ID']}.pdf")
            send_payslip_email(sender_email, app_password, row['EMAIL'], payslip_path)

    except Exception as e:
        print(f"Fatal error: {e}")

if __name__ == "__main__":
    main()
