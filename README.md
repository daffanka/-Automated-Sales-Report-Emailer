# 📧 Automated Sales Report Emailer

This Python script automates the process of sending revenue, bookings, and backlogs report to each salesperson in your organization, broken down per office.

---

## 📌 Features

- Pulls **monthly** and **YTD** data from SQL Server
- Matches salesperson data with email mapping from Excel
- Creates styled **HTML tables** for email body
- Embeds **company header and footer images**
- Sends personalized emails **per salesperson per office**
- Logs email sending status to SQL Server log table

---

## 🧩 Workflow Diagram

```text
Data Collecting (SQL Server)
        ↓
Connect with Email Mapping (Excel)
        ↓
Create HTML Tables (Monthly & YTD)
        ↓
Send Email per Salesperson per Office
        ↓
Log Status to SQL Server
```

---

## 🔧 Requirements

```bash
pip install pandas pyodbc premailer beautifulsoup4
```

---

## 🗂️ Folder Structure

```
project/
│
├── ssms_to_email_rev_report_sp_v3.py      # Main script
├── email_receiver.xlsx                    # Excel mapping: Username & Email
├── email header.png                       # Header image for email
├── footer email.png                       # Footer image for email
```

---

## 🧠 Tech Stack

| Tool               | Purpose                              |
|--------------------|---------------------------------------|
| `pandas`           | Read Excel, manipulate tabular data  |
| `pyodbc`           | SQL Server data connection           |
| `smtplib`          | Send email via Office365             |
| `email.mime`       | Format HTML + Image email            |
| `premailer`        | Inline CSS styling for email HTML    |
| `BeautifulSoup`    | Fix raw HTML table to be email-safe  |

---

## 📤 Output Sample

Email sent will include:
- Company header
- Personalized greeting
- YTD Summary Table
- Monthly Summary Table
- Footer with image

---

## 📥 SQL Tables Used

- `view_weekly_salesperson_rep_monthly`
- `view_weekly_salesperson_rep_ytd`
- `email_log_salesperson_report` (log)

---

## 📌 How to Run

```bash
python email_sender.py
```

---

## 📈 Next Improvements

- Email preview before sending ✅
- Add retry logic for failed emails 🔄
- Parameterize config (YAML/ENV) 🔧
- Add attachment support 📎
- Schedule via Task Scheduler ⏰

---

## 👨‍💻 Author

Muhammad Nauval Daffanka  
Business Intelligence Analyst, Esco Lifesciences  

---

## 🔒 Notes

Do not push credentials to public repo. Use `.env` or secret manager for production.
