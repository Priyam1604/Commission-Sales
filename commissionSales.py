import os
import tempfile
import pyodbc
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from configparser import ConfigParser
from tkinter import simpledialog
from tkinter import messagebox
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import xlsxwriter

def save_config(server, database, username, password, authentication, smtp_server, smtp_port, email_sender, email_password, email_recipients):
    config = ConfigParser()

    if 'SQL_SERVER' not in config.sections():
        config.add_section('SQL_SERVER')
    config.set('SQL_SERVER', 'server', server)
    config.set('SQL_SERVER', 'database', database)
    config.set('SQL_SERVER', 'username', username)
    config.set('SQL_SERVER', 'password', password)
    config.set('SQL_SERVER', 'authentication', authentication)

    if 'SMTP' not in config.sections():
        config.add_section('SMTP')
    config.set('SMTP', 'smtp_server', smtp_server)
    config.set('SMTP', 'smtp_port', smtp_port)
    config.set('SMTP', 'email_sender', email_sender)
    config.set('SMTP', 'email_password', email_password)
    config.set('SMTP', 'email_recipients', ','.join(email_recipients))

    with open('config.ini', 'w') as configfile:
        config.write(configfile)

def read_config():
    config = ConfigParser()
    config.read('config.ini')
    
    server = config.get('SQL_SERVER', 'server')
    database = config.get('SQL_SERVER', 'database')
    username = config.get('SQL_SERVER', 'username')
    password = config.get('SQL_SERVER', 'password')
    authentication = config.get('SQL_SERVER', 'authentication')

    smtp_server = config.get('SMTP', 'smtp_server')
    smtp_port = config.getint('SMTP', 'smtp_port')
    email_sender = config.get('SMTP', 'email_sender')
    email_password = config.get('SMTP', 'email_password')
    email_recipients = config.get('SMTP', 'email_recipients').split(',')

    return server, database, username, password, authentication, smtp_server, smtp_port, email_sender, email_password, email_recipients

def send_email(smtp_server, smtp_port, email_sender, email_password, email_recipients, file_path):
    msg = MIMEMultipart()
    msg['From'] = email_sender
    msg['To'] = ', '.join(email_recipients)
    msg['Subject'] = 'Weekly Report'

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(file_path, "rb").read())
    encoders.encode_base64(part)

    part.add_header('Content-Disposition', 'attachment', filename='report.csv')  # use a generic filename
    msg.attach(part)

    smtp = smtplib.SMTP(smtp_server, smtp_port)
    smtp.starttls()
    smtp.login(email_sender, email_password)
    smtp.sendmail(email_sender, email_recipients, msg.as_string())
    smtp.quit()

def create_conn_string(server, database, username, password, authentication):
    if authentication == 'SQL':
        return f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
    else:
        return f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes;'



def generate_report(email_report=False):
    server, database, username, password, authentication, smtp_server, smtp_port, email_sender, email_password, email_recipients = read_config()
    conn_string = create_conn_string(server, database, username, password, authentication)
    conn = pyodbc.connect(conn_string)

    query = """
            SELECT 
    b.Name AS Store,
    c.LastName AS Supplier,
    i.Supplier AS "Group",
    i.UPC AS PLU,
    i.Description AS "Product",
    i.Field_Integer AS Rate,
    SUM(tl.SubAfterTax) AS Price,   
    SUM(CASE WHEN DATEPART(dw, th.Logged) = 2 THEN tl.Quantity ELSE 0 END) AS MON_QTY,
    SUM(CASE WHEN DATEPART(dw, th.Logged) = 3 THEN tl.Quantity ELSE 0 END) AS TUE_QTY,
    SUM(CASE WHEN DATEPART(dw, th.Logged) = 4 THEN tl.Quantity ELSE 0 END) AS WED_QTY,
    SUM(CASE WHEN DATEPART(dw, th.Logged) = 5 THEN tl.Quantity ELSE 0 END) AS THU_QTY,
    SUM(CASE WHEN DATEPART(dw, th.Logged) = 6 THEN tl.Quantity ELSE 0 END) AS FRI_QTY,
    SUM(CASE WHEN DATEPART(dw, th.Logged) = 7 THEN tl.Quantity ELSE 0 END) AS SAT_QTY,
    SUM(CASE WHEN DATEPART(dw, th.Logged) = 1 THEN tl.Quantity ELSE 0 END) AS SUN_QTY,
    SUM(CASE WHEN DATEPART(dw, th.Logged) = 2 THEN tl.SubAfterTax ELSE 0 END) AS Mon_Sales,
    SUM(CASE WHEN DATEPART(dw, th.Logged) = 3 THEN tl.SubAfterTax ELSE 0 END) AS Tue_Sales,
    SUM(CASE WHEN DATEPART(dw, th.Logged) = 4 THEN tl.SubAfterTax ELSE 0 END) AS Wed_Sales,
    SUM(CASE WHEN DATEPART(dw, th.Logged) = 5 THEN tl.SubAfterTax ELSE 0 END) AS Thu_Sales,
    SUM(CASE WHEN DATEPART(dw, th.Logged) = 6 THEN tl.SubAfterTax ELSE 0 END) AS Fri_Sales,
    SUM(CASE WHEN DATEPART(dw, th.Logged) = 7 THEN tl.SubAfterTax ELSE 0 END) AS Sat_Sales,
    SUM(CASE WHEN DATEPART(dw, th.Logged) = 1 THEN tl.SubAfterTax ELSE 0 END) AS Sun_Sales
FROM 
    AKPOS.dbo.Items i 
    JOIN AKPOS.dbo.TransLines tl ON i.UPC = tl.UPC 
    JOIN AKPOS.dbo.TransHeaders th ON tl.TransNo = th.TransNo AND tl.Branch = th.Branch AND tl.Station = th.Station
    JOIN AKPOS.dbo.Branches b ON th.Branch = b.ID
    JOIN AKPOS.dbo.Customers c ON th.Customer = c.Code
WHERE th.Logged BETWEEN '2020-10-03' AND '2022-10-09' and Field_Integer != 0
GROUP BY 
    b.Name, 
    c.LastName, 
    i.Supplier,
    i.UPC,
    i.Description,
    i.Field_Integer
        """

    df = pd.read_sql(query, conn)
    conn.close()

    # Calculate the "Sales Total" column as the sum of daily sales
    df['Sales Total'] = df[['Mon_Sales', 'Tue_Sales', 'Wed_Sales', 'Thu_Sales', 'Fri_Sales', 'Sat_Sales', 'Sun_Sales']].sum(axis=1)

    # Calculate Commission and Net
    df['Commission'] = df['Sales Total'] * df['Rate'] / 100
    df['Net'] = df['Sales Total'] - df['Commission']

    # NEW CODE BLOCK START
    df_final = pd.DataFrame()

    unique_combinations = df[['Store', 'Group', 'Supplier']].drop_duplicates()

    for _, row in unique_combinations.iterrows():
        store, group, supplier = row['Store'], row['Group'], row['Supplier']

        subset = df[(df['Store'] == store) & (df['Group'] == group) & (df['Supplier'] == supplier)]

        totals = subset[['Mon_Sales', 'Tue_Sales', 'Wed_Sales', 'Thu_Sales', 'Fri_Sales', 'Sat_Sales', 'Sun_Sales', 'Sales Total', 'Commission', 'Net']].sum()
        totals['Supplier'] = supplier + ' TOTAL'
        totals['Store'] = ''
        totals['Group'] = ''
        totals['PLU'] = ''
        totals['Product'] = ''
        totals['Rate'] = ''

        subset = pd.concat([subset, pd.DataFrame(totals).transpose()], ignore_index=True)
        df_final = pd.concat([df_final, subset], ignore_index=True)

    df_final.columns = pd.MultiIndex.from_tuples([
        ('', 'Store'), 
        ('', 'Supplier'), 
        ('', 'Group'),
        ('', 'PLU'),
        ('', 'Product'),
        ('', 'Rate'),
        (' ', 'Price'),
        (' ', 'MON_QTY'),
        (' ', 'TUE_QTY'),
        (' ', 'WED_QTY'),
        (' ', 'THU_QTY'),
        (' ', 'FRI_QTY'),
        (' ', 'SAT_QTY'),
        (' ', 'SUN_QTY'),
        ('Sales', 'Mon_Sales'),
        ('Sales', 'Tue_Sales'),
        ('Sales', 'Wed_Sales'),
        ('Sales', 'Thu_Sales'),
        ('Sales', 'Fri_Sales'),
        ('Sales', 'Sat_Sales'),
        ('Sales', 'Sun_Sales'),
        ('Commission Calculation', 'Sales Total'),
        ('Commission Calculation', 'Commission'),
        ('Commission Calculation', 'Net'),
    ])

    

    # NEW CODE BLOCK END

    # Adding '$' to all necessary fields
    dollar_fields = [
        ('Sales', 'Mon_Sales'),
        ('Sales', 'Tue_Sales'),
        ('Sales', 'Wed_Sales'),
        ('Sales', 'Thu_Sales'),
        ('Sales', 'Fri_Sales'),
        ('Sales', 'Sat_Sales'),
        ('Sales', 'Sun_Sales'),
        ('Commission Calculation', 'Sales Total'),
        ('Commission Calculation', 'Commission'),
        ('Commission Calculation', 'Net')
    ]

    for field in dollar_fields:
        df_final[field] = df_final[field].apply(lambda x: f'${x:.2f}')
    df_final['V3'] = '=SUM(O3:U3)'

    df_single_level_cols = df_final.copy()
    df_single_level_cols.columns = ['_'.join(col) for col in df_final.columns]
    
    if email_report:
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
            file_path = f.name
            
        writer = pd.ExcelWriter(file_path, engine='xlsxwriter')

        df_single_level_cols.to_excel(writer, sheet_name='Sheet1', index=False, startrow=1) # Added startrow=1

        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']

        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })

        for col_num, value in enumerate(df_single_level_cols.columns.values):
            worksheet.write(1, col_num, value, header_format) 

        worksheet.merge_range('O1:U1', 'Sales', header_format)
        worksheet.merge_range('V1:X1', 'Commission Calculation', header_format)

        writer.close()

        send_email(smtp_server, smtp_port, email_sender, email_password, email_recipients, file_path)
        os.unlink(file_path)  # delete file after sending the email

    else:
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if file_path:
            writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
            df_single_level_cols.to_excel(writer, sheet_name='Sheet1', index=False, startrow=1) # Added startrow=1

            workbook  = writer.book
            worksheet = writer.sheets['Sheet1']

            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1
            })

            for col_num, value in enumerate(df_single_level_cols.columns.values):
                worksheet.write(1, col_num, value, header_format)

            worksheet.merge_range('O1:U1', 'Sales', header_format)
            worksheet.merge_range('V1:Y1', 'Commission Calculation', header_format)

            writer.close()
            
            messagebox.showinfo('Success', 'Report generated successfully!')




def enter_credentials():
    dialog = tk.Toplevel(root)
    dialog.title("Database and Email Credentials")
    inputs = ['server', 'database', 'username', 'password', 'authentication', 'smtp_server', 'smtp_port', 'email_sender', 'email_password', 'email_recipients']

    entries = {}

    for i, input in enumerate(inputs):
        label = tk.Label(dialog, text=input.capitalize())
        label.grid(row=i, column=0)
        entry = tk.Entry(dialog)
        if input in ['password', 'email_password']:
            entry.config(show='*')
        elif input == 'email_recipients':
            entry.insert(0, "email1@example.com,email2@example.com")
        entry.grid(row=i, column=1)
        entries[input] = entry

    def submit():
        server = entries['server'].get()
        database = entries['database'].get()
        username = entries['username'].get()
        password = entries['password'].get()
        authentication = entries['authentication'].get()
        smtp_server = entries['smtp_server'].get()
        smtp_port = entries['smtp_port'].get()
        email_sender = entries['email_sender'].get()
        email_password = entries['email_password'].get()
        email_recipients = entries['email_recipients'].get().split(',')
        
        save_config(server, database, username, password, authentication, smtp_server, smtp_port, email_sender, email_password, email_recipients)
        dialog.destroy()

    button_submit = tk.Button(dialog, text="Submit", command=submit)
    button_submit.grid(row=len(inputs), column=0, columnspan=2)


root = tk.Tk()
root.geometry('200x200')
button_config = tk.Button(root, text="Enter Database and Email Credentials", command=enter_credentials)
button_config.pack()
button_generate = tk.Button(root, text="Generate Report", command=lambda: generate_report(False))
button_generate.pack()

# Generate and email report automatically when program starts
generate_report(True)

root.mainloop()