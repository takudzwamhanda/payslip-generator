
import pandas as pd
from fpdf import FPDF
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import ssl
import os
import random
from datetime import datetime, timedelta

def generate_random_phone():
    return f"+44 7{random.randint(100,999)} {random.randint(100,999)} {random.randint(1000,9999)}"

def main():
    try:
        # Debug: Check directory permissions
        payslips_dir = 'payslips'
        if not os.path.exists(payslips_dir):
            os.makedirs(payslips_dir)
            print(f"Created directory: {payslips_dir}")
        
        if not os.access(payslips_dir, os.W_OK):
            print("Warning: Directory is not writable - check permissions")
            return
        
        # Debug: Check Excel file
        print("Reading Excel file...")
        df = pd.read_excel('employees.xlsx.xlsx')
        print("\nFirst few rows of data:")
        print(df.head())
        print("\nData types:")
        print(df.dtypes)
        
        # Clean up and calculate Net Salary
        df['BASIC SALARY'] = df['BASIC SALARY'].replace(',', '', regex=True).astype(float)
        df['ALLOWANCES'] = df['ALLOWANCES'].replace(',', '', regex=True).astype(float)
        df['DEDUCTIONS'] = df['DEDUCTIONS'].replace(',', '', regex=True).astype(float)
        df['Net Salary'] = df['BASIC SALARY'] + df['ALLOWANCES'] - df['DEDUCTIONS']
        
        def generate_payslip(employee):
            try:
                print(f"Generating payslip for {employee['NAME']}...")
                
                # Create PDF with modern styling
                pdf = FPDF(format='A4')
                pdf.add_page()
                
                # Add company logo with modern styling
                try:
                    pdf.image('company_logo.png', x=50, y=10, w=100)
                except:
                    print("Warning: Company logo not found")
                
                # Modern header with gradient effect
                pdf.set_line_width(0.5)
                pdf.set_fill_color(255, 165, 0)  # Orange
                pdf.rect(10, 60, 190, 10, style='F')
                pdf.set_text_color(255, 255, 255)  # White text
                pdf.set_font('Arial', 'B', 16)
                pdf.cell(200, 10, 'AlvisTacool Automotive', ln=True, align='C')
                
                # Reset colors for content
                pdf.set_text_color(0, 0, 0)  # Black text
                pdf.set_font('Arial', 'I', 12)
                pdf.cell(200, 10, 'Payroll Department', ln=True, align='C')
                pdf.ln(15)
                
                # Company Contact Information with modern layout
                pdf.set_font('Arial', '', 10)
                company_info = [
                    'Address: 123 Main Street, London, UK',
                    'Phone: ' + generate_random_phone(),
                    'Email: payroll@alvistacool.com'
                ]
                for info in company_info:
                    pdf.cell(200, 10, info, ln=True, align='L')
                pdf.ln(10)
                
                # Employee Information Section with modern styling
                pdf.set_font('Arial', 'B', 14)
                pdf.cell(200, 10, 'Employee Payslip', ln=True, align='C')
                
                # Employee Details with modern layout
                pdf.set_font('Arial', '', 12)
                details = [
                    ('Employee Name:', employee['NAME']),
                    ('Employee Number:', employee['EMPLOYEE ID']),
                    ('Pay Period:', pd.Timestamp.now().strftime('%d-%B-%Y'))
                ]
                
                x = 15  # Left margin
                y = 100
                for key, value in details:
                    pdf.set_xy(x, y)
                    pdf.cell(100, 10, f"{key} {value}", align='L')
                    y += 10
                    if y > 250:
                        y = pdf.get_y()
                        x = 120
        
                # Salary Breakdown Section with modern styling
                pdf.ln(20)
                pdf.set_font('Arial', 'B', 14)
                pdf.cell(200, 10, 'Salary Breakdown:', ln=True, align='L')
                
                # Modern decorative separator
                pdf.set_line_width(0.5)
                pdf.set_draw_color(255, 165, 0)  # Orange
                pdf.line(10, pdf.get_y(), 200, pdf.get_y())
                
                # Format salary numbers with commas
                net_salary = f"{employee['Net Salary']:,.2f}"
                basic_salary = f"{employee['BASIC SALARY']:,.2f}"
                allowances = f"{employee['ALLOWANCES']:,.2f}"
                deductions = f"{employee['DEDUCTIONS']:,.2f}"
                
                # Salary details with improved layout
                salary_details = [
                    ('Basic Salary:', basic_salary),
                    ('Allowances:', allowances),
                    ('Total Deductions:', deductions),
                    ('Net Salary:', net_salary)
                ]
                
                x = 15  # Left margin
                y = pdf.get_y() + 10
                for key, value in salary_details:
                    pdf.set_xy(x, y)
                    pdf.cell(100, 10, f"{key} {value}", align='L')
                    y += 10
                    if y > 250:
                        y = pdf.get_y()
                        x = 120
        
                # Modern footer with gradient effect
                pdf.ln(20)
                pdf.set_line_width(0.5)
                pdf.set_fill_color(255, 165, 0)  # Orange
                pdf.rect(10, 260, 190, 10, style='F')
                pdf.set_text_color(255, 255, 255)  # White text
                pdf.set_font('Arial', 'I', 10)
                pdf.cell(200, 10, 'This payslip is computer generated and does not require a signature.', ln=True, align='C')
                
                # Save PDF
                pdf_path = f"Payslip_{employee['EMPLOYEE ID']}.pdf"
                pdf.output(pdf_path)
                print(f"PDF saved to: {pdf_path}")
                
            except Exception as e:
                print(f"Error generating payslip: {str(e)}")
        
        # Generate payslips
        for _, row in df.iterrows():
            generate_payslip(row)
        
        # Send emails
        def send_payslip_email(sender_email, app_password, recipient_email, payslip_path):
            try:
                context = ssl.create_default_context()
                server = smtplib.SMTP('smtp.gmail.com', 587)
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
                    attach.add_header('Content-Disposition', 'attachment', filename=payslip_path.split('/')[-1])
                    message.attach(attach)
                
                server.send_message(message)
                print(f"Email sent successfully to {recipient_email}")
                
            except Exception as e:
                print(f"Error sending email to {recipient_email}: {str(e)}")
            finally:
                server.quit()
        
        # Send emails to employees
        sender_email = "mhandatakudzwa07@gmail.com"
        app_password = "asro bhjv qaup gpem"  # Replace with actual app password
        
        for _, row in df.iterrows():
            payslip_path = f"Payslip_{row['EMPLOYEE ID']}.pdf"
            send_payslip_email(sender_email, app_password, row['EMAIL'], payslip_path)
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()