import pandas as pd
import smtplib
import os
import logging
from dotenv import load_dotenv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import time
import random

# Load environment variables
load_dotenv()

# Set up logging
logging.basicConfig(filename="email_log.log", level=logging.INFO)

# Load CSV file
excel_file = "your csv file.csv"
if not os.path.exists(excel_file):
    raise FileNotFoundError(f"CSV file not found: {excel_file}")

df = pd.read_csv(excel_file)
if df.empty:
    raise ValueError("CSV file is empty")

# Get SMTP credentials
smtp_username = os.getenv("SMTP_USERNAME")
smtp_password = os.getenv("SMTP_PASSWORD")
if not smtp_username or not smtp_password:
    raise ValueError("SMTP_USERNAME and SMTP_PASSWORD must be set in the .env file")

# Email body template
email_body_template = """
Respected Sir/Ma'am,
<br>
<br>
I'm Tanishq Ranjan, a 4th Year engineering undergrad at the National Institute of Technology Delhi. I came across {company} and was impressed by its remarkable growth. If feasible I'd be excited to explore potential internship opportunities within {company}. 
<br>
<br>
<strong>About Me:</strong>
<br>
<br>
<strong>1. Education:</strong>
<br>
- B.Tech Major in Electronics and Communication Engineering with a Minor in AI & ML from NIT Delhi, Graduating 2025.
<br>
<br>
<strong>2. Experience:</strong>
<br>
- SDE Intern at Mavenir Systems Pvt Ltd
<br>
- Machine Learning Project Intern at DRDO - LRDE
<br>
- Big Data Analytics & ML Intern at DRDO - ISSA
<br>
- AI Project Intern at Ministry of Electronics & IT, Govt. of India
<br>
- Full Stack Web Development Intern at CYRAN AI Solutions, IIT Delhi
<br>
- LeetCode Rating: 1930+ (Knight ranked)
<br>
- Selected for Amazon ML School 2024
<br>
- Secured 98.4th Percentile out of 1.1 million candidates in JEE Mains 2021
<br>
<br>
<strong>3. Technical Skills:</strong>
<br>
- Programming Languages: C++, Python, JavaScript
<br>
- Machine Learning & AI: PyTorch, TensorFlow, Neural Networks, NLP, Computer Vision
<br>
- Web Development: MongoDB, MySQL, Express.js, Node.js, CSS
<br>
- Big Data & DevOps: Apache Kafka, Apache Spark
<br>
- Frameworks & Libraries: Streamlit, JQuery, Scikit-Learn, Pandas, Numpy, Seaborn
<br>
<br>
I would be thrilled to discuss how my background, skills, and projects align with the goals of {company}. If you're available, I would love to schedule a call to explore this opportunity further.
<br>
<br>
Thank you for your time and consideration.
<br>
<br>
Please find my resume attached. You can also connect with me on:
<br>
<a href="{linkedin_link}">LinkedIn</a> | <a href="{github_link}">GitHub</a> | <a href="{portfolio_link}">Leetcode</a>
<br>
<br>
Regards,
<br>
<br>
<strong>Tanishq Ranjan</strong>
<br>
4th Year Undergraduate Student
<br>
Department of Electronics and Communication Engineering
<br>
<strong>National Institute of Technology Delhi</strong>
<br>
<br>
<strong>Email:</strong> 211220058@nitdelhi.ac.in
<br>
<strong>Phone:</strong> (+91) 9113963397
<br>
<br>
====================================================================
<br>
NATIONAL INSTITUTE OF TECHNOLOGY DELHI [NITD]
"""

# Check if PDF exists
pdf_path = "TanishqRanjan.pdf"
if not os.path.exists(pdf_path):
    raise FileNotFoundError(f"PDF file not found: {pdf_path}")

# Send emails
try:
    with smtplib.SMTP("smtp.gmail.com:587") as server:
        server.ehlo()
        server.starttls()
        server.login(smtp_username, smtp_password)

        for index, row in df.iterrows():
            time.sleep(random.randint(5, 15))  # Increased delay to avoid spam flags

            message = MIMEMultipart()
            message["Subject"] = (
                f"Internship Inquiry cum Application @{row['Company']} - Tanishq Ranjan, NIT Delhi"
            )
            message["From"] = smtp_username
            message["To"] = row["Email"]

            email_body = email_body_template.format(
                name=row["First Name"],
                company=row["Company"],
                linkedin_link="https://in.linkedin.com/in/tanishq-ranjan-26a52740",
                github_link="https://github.com/x-INFiN1TY-x",
                portfolio_link="https://leetcode.com/u/211220058/",
            )
            message.attach(MIMEText(email_body, "html"))

            with open(pdf_path, "rb") as pdf_file:
                pdf_attachment = MIMEApplication(pdf_file.read(), _subtype="pdf")
                pdf_attachment.add_header(
                    "Content-Disposition", f"attachment; filename={pdf_path}"
                )
                message.attach(pdf_attachment)

            try:
                server.sendmail(smtp_username, row["Email"], message.as_string())
                logging.info(f"Email sent to {row['Email']}")
            except Exception as e:
                logging.error(f"Error sending email to {row['Email']}: {e}")

except Exception as e:
    logging.error(f"Error in email sending process: {e}")

logging.info("Email sending process completed")
