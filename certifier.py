import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import arabic_reshaper
from bidi.algorithm import get_display
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import webbrowser
#
# Function to select Excel file
def select_excel():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if filepath:
        excel_path.set(filepath)

# Function to send a test email 
def send_test_email():
    messagebox.showinfo("Test Email", "Test email sent!")

# Function to send bulk emails 
def send_emails():
    messagebox.showinfo("Bulk Emails", "Bulk emails sent!")

# Load the Excel file
df = pd.read_excel('sheet.xlsx') #the path to ur excel sheet file (or u can rename it to this)
# also make sure that ur excel file have each row labeled with Name, Phone, Email, to avoid errors.
# The code will automatically look for these names to access the data correctly

template_path = 'template.jpeg' #specify the template's path
cordsX = 486 # the X of the place where the name should be printed
cordsY = 342 # the Y of the place where the name should be printed
# You can use gimp (or Paint if u r on windows) to find those coordinates
# The text is centered, so the point should be centered on the space u desire to have the names on

certFont = 'DejaVuSans-Bold.ttf' # Here specify the font path
fontSize = 24
subject = "Your certificate" # The subject of the email
body = """ 
Congrats
""" # The body of the email (Can have multiple lines)
urEmail = 'email here' # The email you are using for sending
urPass = 'password here' # The password of that email

# Certificate generation function
def generate_certificate(name, email, template_path=template_path, output_path='certificates/'):
    # Create output directory if it doesn't exist
    if not os.path.exists(output_path):
        os.makedirs(output_path)
    
    # Load the certificate template
    image = Image.open(template_path)
    draw = ImageDraw.Draw(image)
    
    # Define the font and size 
    font_path = certFont
    font_size = fontSize
    try:
        font = ImageFont.truetype(font_path, font_size)
    except IOError:
        raise Exception(f"Font file not found: {font_path}")
    font = ImageFont.truetype(font_path, font_size)
    
    # Target coordinates (center of the text)
    target_x = cordsX
    target_y = cordsY

    # Ensure the text fits within the designated area
    max_width = 1000  # Maximum width for the text
    
    # Resize the font if necessary to fit within the max_width
    bbox = draw.textbbox((0, 0), name, font=font)
    textwidth = bbox[2] - bbox[0]
    textheight = bbox[3] - bbox[1]
    while textwidth > max_width:
        font_size -= 1
        font = ImageFont.truetype(font_path, font_size)
        bbox = draw.textbbox((0, 0), name, font=font)
        textwidth = bbox[2] - bbox[0]
        textheight = bbox[3] - bbox[1]



    # Calculate the position to center the text at the target coordinates
    x = target_x - (textwidth / 2)
    y = target_y - (textheight / 2)

    # Draw the name on the certificate
    reshaped_text = arabic_reshaper.reshape(name)
    bidi_text = get_display(reshaped_text)
    draw.text((x, y), name, font=font, fill="black")  
    
    # Save the certificate
    output_file = output_path + name.replace(" ", "_") + '.png'
    image.save(output_file)
    
    return output_file

# Email sending function
def send_email(to_email, subject, body, attachment_path, from_email=urEmail, password=urPass): 
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
        certificate_path = generate_certificate(name, email)
        send_email(email, subject, body, certificate_path) #(email, subject, body, path)
    except Exception as e:
        log_error(name, phone, str(e))



root = tk.Tk()
root.title("Certifier - by oAmadu on github")

excel_path = tk.StringVar()
tk.Label(root, text="Select Excel File:").pack()
tk.Entry(root, textvariable=excel_path).pack()
tk.Button(root, text="Browse", command=select_excel).pack()

urEmail = tk.StringVar()
emailFrame = tk.Frame(root) 
emailFrame.pack()
tk.Label(emailFrame, text="Your Email").pack(side=tk.LEFT)
tk.Entry(emailFrame, textvariable=urEmail).pack(side=tk.BOTTOM)

def show_explanation():
    explanation_window = tk.Toplevel(root)
    explanation_window.title("How to get the generated password")
    explanation_text = tk.Label(explanation_window, text="To generate an app-specific password:\n\n1. Go to your email account settings.\n2. Navigate to the security section.\n3. Generate an app-specific password.\n4. Copy and paste it here.  \n\n  For more details, visit:", padx=10, pady=10, wraplength=400)
    explanation_text.pack(side=tk.TOP)
    link = tk.Label(explanation_window, text="Google App Passwords", fg="blue", cursor="hand2")
    link.pack()
    link.bind("<Button-1>", lambda e: webbrowser.open_new("https://myaccount.google.com/apppasswords"))

#Password Entry section
urPass = tk.StringVar()
frame = tk.Frame(root)
frame.pack()

tk.Label(frame, text="Your generated password").pack(side=tk.LEFT)

# Add "How do I get that?" clickable text
howtoget = tk.Label(frame, text="(How do I get that?)", fg="blue", cursor="hand2")
howtoget.pack(side=tk.RIGHT)
howtoget.bind("<Button-1>", lambda e: show_explanation())

passEntryFrame = tk.Frame(frame)
passEntryFrame.pack()

tk.Entry(passEntryFrame, textvariable=urPass).pack(side=tk.LEFT)



tk.Button(root, text="Send Test Email", command=send_test_email).pack()
tk.Button(root, text="Send Bulk Emails", command=send_emails).pack()



root.mainloop()
