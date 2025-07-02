import smtplib
import pandas as pd
import pyodbc
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from premailer import transform
from datetime import datetime
from bs4 import BeautifulSoup

# --- Config ---
email_map_path = r"C:\Users\muhammad.nauval\OneDrive - Esco (1)\ESCO-BI - General\email_receiver.xlsx"
server = '203.127.53.9'
db_data = 'BCDB'
db_log = 'SYSDB'
sender_email = "muhammad.nauval@escolifesciences.com"
password = "@Escodaffa1"
tanggal_today = datetime.now().strftime("%Y-%m-%d")

# --- CSS Table ---
css_style = """
<style>
    .styled-table {
        border-collapse: collapse;
        font-size: 14px;
        font-family: 'Segoe UI', sans-serif;
        min-width: 500px;
    }
    .styled-table th, .styled-table td {
        padding: 8px 12px;
        border: 1px solid #dddddd;
        text-align: left;
        background-color: white;
        color: black;
    }
    .styled-table tbody tr:nth-child(even) {
        background-color: #f9f9f9;
    }
</style>
"""

# --- Koneksi DB ---
conn_data = pyodbc.connect(f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={db_data};Trusted_Connection=yes;")
conn_log = pyodbc.connect(f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={db_log};Trusted_Connection=yes;")
cursor_log = conn_log.cursor()

# --- Load email mapping ---
df_email_map = pd.read_excel(email_map_path)
df_email_map['Username'] = df_email_map['Username'].str.strip()
nama_sql_string = ", ".join([f"'{nama}'" for nama in df_email_map['Username']])

# --- Load Monthly + YTD ---
query_monthly = f"""
SELECT 
    [Month],
    [MonthNumber],
    [Salesperson_Name],
    [Office],
    [Revenue (SGD)],
    [Revenue (LCY)],
    [Bookings (SGD)],
    [Bookings (LCY)]
FROM [dbo].[view_weekly_salesperson_rep_monthly]
WHERE LTRIM(RTRIM([Salesperson_Name])) IN ({nama_sql_string})
"""

query_ytd = f"""
SELECT 
    [Salesperson_Name],
    [Office],
    [Revenue YTD (SGD)],
    [Revenue YTD (LCY)],
    [Bookings YTD (SGD)],
    [Bookings YTD (LCY)],
    [Backlogs YTD (SGD)],
    [Backlogs YTD (LCY)]
FROM [dbo].[view_weekly_salesperson_rep_ytd]
WHERE LTRIM(RTRIM([Salesperson_Name])) IN ({nama_sql_string})
"""

df_monthly = pd.read_sql(query_monthly, conn_data)
df_ytd = pd.read_sql(query_ytd, conn_data)

# --- Format angka ---
def format_df(df):
    df = df.copy()
    for col in df.select_dtypes(include=['float64', 'int64']).columns:
        df[col] = df[col].apply(lambda x: f"{x:,.0f}" if x != 0 else "0")
    return df

# --- Convert to HTML Table ---
def df_to_html(df):
    html = df.to_html(index=False, classes="styled-table", border=0)
    soup = BeautifulSoup(html, "html.parser")
    return str(soup)

# --- Log email ke DB ---
def log_email(username, email_to, status, error_message=None):
    check_q = """
        SELECT 1 FROM email_log_salesperson_report
        WHERE username = ? AND CAST(sent_at AS DATE) = CAST(GETDATE() AS DATE)
    """
    cursor_log.execute(check_q, username)
    if not cursor_log.fetchone():
        insert_q = """
            INSERT INTO email_log_salesperson_report (username, email, status, error_message, sent_at)
            VALUES (?, ?, ?, ?, GETDATE())
        """
        cursor_log.execute(insert_q, username, email_to, status, error_message)
        conn_log.commit()

# --- Kirim Email ---
def send_email(username, email_to, office, df_office_monthly, df_office_ytd):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = email_to
    msg['Subject'] = f"Revenue, Bookings & Backlogs Report - {office}"

    html_sections = []

    if not df_office_ytd.empty:
        df_ytd_trim = df_office_ytd.drop(columns=['Salesperson_Name', 'Office'], errors='ignore')
        html_sections.append("<h3>YTD Summary</h3>" + df_to_html(format_df(df_ytd_trim)))

    if not df_office_monthly.empty:
        df_month_trim = df_office_monthly[['Month', 'Revenue (SGD)', 'Revenue (LCY)', 'Bookings (SGD)', 'Bookings (LCY)']]
        html_sections.append("<h3>Monthly Summary</h3>" + df_to_html(format_df(df_month_trim)))

    email_body = f"""
    <html>
    <head>{css_style}</head>
    <body>
        <div><img src="cid:headerimg" style="width:100%; max-width:700px;"></div>
        <p>Dear {username},</p>
        <p>As of {tanggal_today}, your revenue, backlogs, and bookings for <b>{office}</b> are as follows:</p>
        {''.join(html_sections)}
        <p>Regards,<br>BI Analyst Team</p>
        <div><img src="cid:footerimg" style="width:100%; max-width:700px; margin-top: 20px;"></div>
    </body>
    </html>
    """

    email_body_inlined = transform(email_body)
    msg.attach(MIMEText(email_body_inlined, 'html'))

    def attach_image(path, cid):
        with open(path, 'rb') as f:
            img = MIMEImage(f.read())
            img.add_header('Content-ID', f'<{cid}>')
            img.add_header('Content-Disposition', 'inline', filename=os.path.basename(path))
            msg.attach(img)

    attach_image(r"C:\Users\muhammad.nauval\OneDrive - Esco (1)\ESCO-BI - General\email header.png", "headerimg")
    attach_image(r"C:\Users\muhammad.nauval\OneDrive - Esco (1)\ESCO-BI - General\footer email.png", "footerimg")

    try:
        with smtplib.SMTP("smtp.office365.com", 587) as smtp:
            smtp.starttls()
            smtp.login(sender_email, password)
            smtp.send_message(msg)
        print(f"(✓) Email terkirim ke {username} - {office}")
        log_email(username, email_to, "Success")
    except Exception as e:
        print(f"(-) Gagal kirim ke {username} - {office}: {e}")
        log_email(username, email_to, "Failed", str(e))

# --- Looping kirim ---
for _, row in df_email_map.iterrows():
    username = row['Username'].strip()
    email_to = row['Email'].strip()

    df_user_monthly = df_monthly[df_monthly['Salesperson_Name'].str.strip().str.lower() == username.lower()]
    df_user_ytd = df_ytd[df_ytd['Salesperson_Name'].str.strip().str.lower() == username.lower()]

    if df_user_monthly.empty and df_user_ytd.empty:
        continue

    all_offices = pd.concat([df_user_monthly['Office'], df_user_ytd['Office']]).dropna().unique()
    for office in all_offices:
        df_office_monthly = df_user_monthly[df_user_monthly['Office'] == office].sort_values('MonthNumber')
        df_office_ytd = df_user_ytd[df_user_ytd['Office'] == office]
        send_email(username, email_to, office, df_office_monthly, df_office_ytd)
