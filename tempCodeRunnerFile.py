#pip install pandas pillow openpyxl

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# Load the Excel file
df = pd.read_excel('sheet.xlsx')

subject = "Your certificate"
body = """
Congrats
"""
urEmail = 'aymt7mitests@gmail.com'
urPass = 'c'

# Certificate generation function
def generate_certificate(name, template_path='template.jpeg', output_path='certificates/'):
    # Create output directory if it doesn't exist
    if not os.path.exists(output_path):
        os.makedirs(output_path)
    
    # Load the certificate template
    image = Image.open(template_path)
    draw = ImageDraw.Draw(image)
    
    # Define the font and size 
    font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"  # font path
    font_size = 52
    font = ImageFont.truetype(font_path, font_size)
    
    # Target coordinates (center of the text)
    target_x = 1001
    target_y = 421

    # Ensure the text fits within the designated area
    max_width = 1000  # Maximum width for the text
    
    # Resize the font if necessary to fit within the max_width
    textwidth, textheight = draw.textsize(name, font)
    while textwidth > max_width:
        font_size -= 1
        font = ImageFont.truetype(font_path, font_size)
        textwidth, textheight = draw.textsize(name, font)

    # Calculate the position to center the text at the target coordinates
    x = target_x - (textwidth / 2)
    y = target_y - (textheight / 2)

    # Draw the name on the certificate
    draw.text((x, y), name, font=font, fill="black")  # Adjust the fill color to match the rest of the text
    
    # Save the certificate
    output_file = output_path + name.replace(" ", "_") + '.png'
    image.save(output_file)
    
    return output_file

# Email sending function
def send_email(to_email, subject, body, attachment_path, from_email=urEmail, password=urPass): # Adjust this with tanmya's credentials.
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    
    attachment = open(attachment_path, 'rb')
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % attachment_path)
    msg.attach(part)
    
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(from_email, password)
    text = msg.as_string()
    server.sendmail(from_email, to_email, text)
    server.quit()

# Logging function
def log_error(name, phone, reason):
    with open('error_log.txt', 'a') as log_file:
        log_file.write(f"Name: {name}, Phone: {phone}, Reason: {reason}\n")

# Main loop to generate certificates and send emails
for index, row in df.iterrows():
    name = row['Name']
    email = row['Email']
    phone = row['Phone']
    
    # Skip if no valid email
    if not isinstance(email, str) or email.lower() in ["no email", ""]:
        log_error(name, phone, "Invalid email")
        continue
    
    try:
        certificate_path = generate_certificate(name)
        send_email(email, subject, body, certificate_path) #(email, subject, body, path)
    except Exception as e:
        log_error(name, phone, str(e))
