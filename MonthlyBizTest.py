import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.cluster import KMeans
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Outlook SMTP configuration
SMTP_SERVER = 'smtp-mail.outlook.com'
SMTP_PORT = 587
SMTP_EMAIL = 'forgeanalytics@outlook.com'
SMTP_PASSWORD = os.getenv("OUTLOOK_PASSWORD") 

# Data loading
def load_data():
    customer_data = pd.read_csv('/Users/mattobrien/Documents/Mock_Customer_Data.csv')
    inventory_data = pd.read_csv('/Users/mattobrien/Documents/Mock_Inventory_Data.csv')
    product_data = pd.read_csv('/Users/mattobrien/Documents/Mock_Product_Data.csv')
    sales_data = pd.read_csv('/Users/mattobrien/Documents/Mock_Sales_Data.csv')
    return customer_data, inventory_data, product_data, sales_data

# Data cleaning
def clean_data(df):
    df = df.drop_duplicates()
    df = df.dropna()
    return df

# Data analysis
def sales_trends(sales_data):
    sales_data['date'] = pd.to_datetime(sales_data['date'])
    sales_data['month'] = sales_data['date'].dt.to_period('M')
    sales_data['sales'] = pd.to_numeric(sales_data['sales'], errors='coerce')
    monthly_sales = sales_data.groupby('month')['sales'].sum().reset_index()
    monthly_sales['month'] = monthly_sales['month'].astype(str)
    plt.figure(figsize=(10, 5))
    sns.lineplot(x='month', y='sales', data=monthly_sales)
    plt.title('Monthly Sales Trends')
    plt.xticks(rotation=45)
    plt.savefig('monthly_sales_trends.png')
    plt.close()

# Report generation
def generate_report(sales_data, customer_data):
    sales_trends(sales_data)
    # Further report generation can be added here

# Email the report
def send_email():
    message = MIMEMultipart()
    message['From'] = SMTP_EMAIL
    message['To'] = 'obrien.p.matthew@gmail.com'  # Set the recipient's email
    message['Subject'] = 'Monthly Business Performance Report'
    
    body = 'Find attached the monthly report and relevant visualizations.'
    message.attach(MIMEText(body, 'plain'))

    # Attach files
    with open('monthly_sales_trends.png', 'rb') as f:
        attach = MIMEBase('application', 'octet-stream')
        attach.set_payload(f.read())
        encoders.encode_base64(attach)
        attach.add_header('Content-Disposition', 'attachment', filename='monthly_sales_trends.png')
        message.attach(attach)

    # Send email

    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()
    server.login(SMTP_EMAIL, SMTP_PASSWORD)
    server.send_message(message)
    server.quit()

# Main function
def main():
    customer_data, inventory_data, product_data, sales_data = load_data()
    customer_data = clean_data(customer_data)
    sales_data = clean_data(sales_data)
    generate_report(sales_data, customer_data)
    send_email()

if __name__ == "__main__":
    main()
