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
import re
from email.utils import formataddr
from datetime import datetime, timedelta

# Load environment variables
load_dotenv()

# Set up logging
logging.basicConfig(filename="email_log.log", level=logging.INFO)

# Load CSV file
excel_file = "delhi.csv"
if not os.path.exists(excel_file):
    raise FileNotFoundError(f"CSV file not found: {excel_file}")

df = pd.read_csv(excel_file)
if df.empty:
    raise ValueError("CSV file is empty")

# Get multiple SMTP credentials
smtp_credentials = [
    (os.getenv("SMTP_USERNAME"), os.getenv("SMTP_PASSWORD")),
    # (os.getenv("SMTP_USERNAME2"), os.getenv("SMTP_PASSWORD2")),
    # (os.getenv("SMTP_USERNAME3"), os.getenv("SMTP_PASSWORD3")),
    # Add more as needed
]

# Ensure all credentials are loaded
smtp_credentials = [(user, pwd) for user, pwd in smtp_credentials if user and pwd]
if not smtp_credentials:
    raise ValueError(
        "At least one set of SMTP credentials must be set in the .env file"
    )

# Define content variations
greetings = [
    "Respected Sir/Ma'am,",
    "Dear Sir/Ma'am,",
    "Greetings,",
]

introductions = [
    "I'm Tanishq Ranjan, a 4th Year engineering undergrad at the National Institute of Technology Delhi.",
    "My name is Tanishq Ranjan, and I'm currently in my 4th year of studying engineering at the National Institute of Technology Delhi.",
    "I am Tanishq Ranjan, a final year engineering student at the National Institute of Technology Delhi.",
    "I hope this email finds you well. I'm Tanishq Ranjan, currently pursuing my final year in engineering at NIT Delhi.",
]

job_hunting_phrases = [
    "I came across {company} while job hunting and was impressed by its remarkable growth.",
    "During my job search, I came across {company} and was struck by its impressive growth.",
    "while on lookout for internship opportunities, I found {company} and was highly impressed.",
    "While researching potential career opportunities, {company} caught my attention due to its growth trajectory.",
    "{company} stood out to me during my internship search due to its reputation for excellence and growth in the industry.",
]

closing_remarks = [
    "I would be thrilled to discuss how my background, skills, and projects align with the goals of {company}.",
    "I am excited about the prospect of discussing how my skills and experiences can contribute to the success of {company}.",
    "I look forward to the opportunity to discuss how my expertise and projects align with the objectives of {company}.",
    "It would be a pleasure to explore how my technical skills and experiences could benefit {company}.",
    "I'm eager to discuss how my passion for technology and my skills could contribute to the innovative work at {company}.",
]

follow_up_phrases = [
    "If you're available, I would love to schedule a call to explore this opportunity further.",
    "Please let me know if we can arrange a call to discuss this potential opportunity in more detail.",
    "I would appreciate the chance to discuss this opportunity further if you have time for a call.",
    "Would it be possible to schedule a brief call to discuss how I could contribute to {company}?",
    "I'm available for a call at your convenience to further discuss this exciting opportunity.",
]

thank_you_phrases = [
    "Thank you for your time and consideration.",
    "Thank you very much for considering my application.",
    "I appreciate your time and look forward to your response.",
    "Thank you for taking the time to review my application.",
]

factual_section = """
<strong>About Me:</strong>
<br>
<strong>1. Education:</strong>
<br>
- B.Tech Major in Electronics and Communication Engineering with a Minor in AI & ML from NIT Delhi, Graduating 2025.
<br>
<br>
<strong>2. Experience:</strong>
<br>
- <u>SDE Intern</u> at <strong>Mavenir Systems Pvt Ltd</strong>
<br>
- <u>Machine Learning Project Intern</u> at <strong>DRDO</strong> - LRDE
<br>
- <u>Big Data Analytics & ML Intern</u> at <strong>DRDO</strong> - ISSA
<br>
- <u.AI Project Intern</u> at <strong>Ministry of Electronics & IT, Govt. of India</strong>
<br>
- <u>Full Stack Web Development Intern</u> at <strong>CYRAN AI Solutions, IIT Delhi</strong>
<br>
- <strong>LeetCode Rating: 1930+ (Knight ranked)</strong>
<br>
- Selected for <strong>Amazon ML School 2024</strong>
<br>
- Secured <strong>98.4th Percentile</strong> out of 1.1 million candidates in JEE Mains 2021
<br>
<strong>3. Technical Skills:</strong>
<br>
- <u>Programming Languages:</u> C++, Python, JavaScript
<br>
- <u>Machine Learning & AI:</u> PyTorch, TensorFlow, Neural Networks, NLP, Computer Vision
<br>
- <u>Web Development:</u> MongoDB, MySQL, Express.js, Node.js, CSS
<br>
- <u>Big Data & DevOps:</u> Apache Kafka, Apache Spark
<br>
- <u>Frameworks & Libraries:</u> Streamlit, JQuery, Scikit-Learn, Pandas, Numpy, Seaborn
<br>
"""

# Check if PDF exists
pdf_path = "TanishqRanjan.pdf"
if not os.path.exists(pdf_path):
    raise FileNotFoundError(f"PDF file not found: {pdf_path}")


def is_valid_email(email):
    # Basic email validation
    regex = "^[a-z0-9]+[\._]?[a-z0-9]+[@]\w+[.]\w+$"
    if re.search(regex, email):
        return True
    return False


def generate_email_content(name, company):
    greeting = random.choice(greetings).format(name=name)
    introduction = random.choice(introductions)
    job_hunting = random.choice(job_hunting_phrases).format(company=company)
    closing_remark = random.choice(closing_remarks).format(company=company)
    follow_up = random.choice(follow_up_phrases).format(company=company)
    thank_you = random.choice(thank_you_phrases).format(company=company)

    email_body = f"""
    {greeting}
    <br><br>
    {introduction}
    {job_hunting} If feasible I'd be excited to explore potential internship opportunities within your org.
    <br><br>
    {factual_section}
    <br><br>
    {closing_remark}
    <br>
    {follow_up}
    <br>
    {thank_you}
    <br><br>
    Please find my resume attached. You can also connect with me on:
    <br>
    <a href="https://in.linkedin.com/in/tanishq-ranjan-26a52740">LinkedIn</a> | <a href="https://github.com/x-INFiN1TY-x">GitHub</a> | <a href="https://leetcode.com/u/211220058/">Leetcode</a>
    <br><br>
    Regards,
    <br><br>
    <strong>Tanishq Ranjan</strong>
    <br>
    4th Year Undergraduate Student
    <br>
    Department of Electronics and Communication Engineering
    <br>
    <strong>National Institute of Technology Delhi [NIT D]</strong>
    <br>
    <strong>Email:</strong> 211220058@nitdelhi.ac.in
    <br>
    <strong>Phone:</strong> +91 9113963397
    """

    return email_body


def send_email(server, from_addr, to_addr, message):
    try:
        server.sendmail(from_addr, to_addr, message.as_string())
        logging.info(f"Email sent to {to_addr}")
        return True
    except Exception as e:
        logging.error(f"Error sending email to {to_addr}: {e}")
        return False


# Rate limiting variables
MAX_DAILY_EMAILS = 50
emails_sent_today = 0
current_date = datetime.now().date()

# Send emails
try:
    for index, row in df.iterrows():
        if not is_valid_email(row["Email"]):
            logging.warning(f"Invalid email address: {row['Email']}")
            continue

        # Check if we've reached the daily limit
        if emails_sent_today >= MAX_DAILY_EMAILS:
            logging.info("Daily email limit reached. Pausing until tomorrow.")
            time.sleep(86400)  # Sleep for 24 hours
            emails_sent_today = 0
            current_date = datetime.now().date()

        # Check if it's a new day
        if datetime.now().date() > current_date:
            emails_sent_today = 0
            current_date = datetime.now().date()

        # Randomly select SMTP credentials
        smtp_username, smtp_password = random.choice(smtp_credentials)

        with smtplib.SMTP("smtp.gmail.com:587") as server:
            server.ehlo()
            server.starttls()
            server.login(smtp_username, smtp_password)

            # Randomize time delay to avoid spam detection
            time.sleep(random.randint(300, 1800) + random.random() * 60)

            message = MIMEMultipart()
            message["Subject"] = (
                f"Internship Inquiry/Application @{row['Company']} - Tanishq Ranjan, NIT Delhi"
            )
            message["From"] = formataddr(("Tanishq Ranjan", smtp_username))
            message["To"] = row["Email"]

            email_body = generate_email_content(row["Name"], row["Company"])
            message.attach(MIMEText(email_body, "html"))

            with open(pdf_path, "rb") as pdf_file:
                pdf_attachment = MIMEApplication(pdf_file.read(), _subtype="pdf")
                pdf_attachment.add_header(
                    "Content-Disposition", f"attachment; filename={pdf_path}"
                )
                message.attach(pdf_attachment)

            if send_email(server, smtp_username, row["Email"], message):
                emails_sent_today += 1

except Exception as e:
    logging.error(f"Error in email sending process: {e}")

logging.info("Email sending process completed")
