# Payslip Generator for AlvisTacool Automotive

A modern Python application for generating professional payslips with email distribution capabilities.

## Features
- Generates professional PDF payslips
- Modern design with orange and black color scheme
- Automatic email distribution
- Error handling and debugging
- Secure file operations
- Random phone number generation

## Prerequisites
- Python 3.7+
- Required packages:
  * pandas
  * fpdf
  * smtplib
  * email
  * ssl
  * os
  * random
  * datetime

## Setup Instructions
1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/payslip-generator.git
   cd payslip-generator
   ```

2. Install required packages:
   ```bash
   pip install pandas fpdf
   ```

3. Create required files:
   - `employees.xlsx` with employee data
   - `company_logo.png` for company branding
   - `.gitignore` file (automatically created)

## Employee Data Format
Your Excel file should contain the following columns:
```markdown
| Column Name    | Description                    |
|---------------|--------------------------------|
| NAME          | Employee full name             |
| EMPLOYEE ID   | Unique employee identifier    |
| EMAIL         | Employee email address        |
| BASIC SALARY  | Monthly basic salary          |
| ALLOWANCES    | Monthly allowances            |
| DEDUCTIONS    | Monthly deductions           |
```

## Configuration
1. Create a `.gitignore` file:
   ```bash
   company_logo.png
   *.pdf
   payslips/
   app_password.txt
   ```

2. Update email configuration:
   - Replace `sender_email` with your Gmail address
   - Replace `app_password` with your Gmail app password

## Running the Script
```bash
python payslips_generator.py
```

## Output
The script will:
1. Generate payslips in the `payslips/` directory
2. Send emails to all employees
3. Display progress and any errors in the console

## Troubleshooting
Common issues:
- File not found: Check Excel file name and path
- Permission errors: Verify write access to directory
- Email errors: Check Gmail app password and sender email
- PDF errors: Verify FPDF installation

## Security Notes
- Keep your app password secure
- Never commit sensitive files to Git
- Use strong passwords for email accounts
- Regularly update dependencies
