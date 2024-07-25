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
from tkinter import font as tkFont
import webbrowser
from PIL import Image, ImageDraw, ImageFont, ImageTk

# Function to select Excel file
def select_excel():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if filepath:
        excel_path.set(filepath)



# Define the global variables for font settings
fontSize = 54
current_font_family = 'Helvetica'
current_font_size = fontSize

# Function to select certificate template
def select_template():
    filepath = filedialog.askopenfilename(filetypes=[("Image files", "*.jpeg;*.png;*.jpg")])
    if filepath:
        cert_path.set(filepath)
        coord_button.config(state=tk.NORMAL)

# Function to open the window for coordinating text place on the certificate
def coordinator():
    global current_font_family, current_font_size
    
    coord_window = tk.Toplevel(root)
    coord_window.title("Select Coordinates")
    
    # Load and display the template image
    template_img = Image.open(cert_path.get())
    template_photo = ImageTk.PhotoImage(template_img)
    canvas = tk.Canvas(coord_window, width=template_img.width, height=template_img.height)
    canvas.pack()
    canvas.create_image(0, 0, anchor=tk.NW, image=template_photo)
    
    # Label to show the coordinates
    coord_label = tk.Label(coord_window, text=f"Coordinates: ({cordsX}, {cordsY})")
    coord_label.pack()
    
    # Add test text to the canvas
    text_id = canvas.create_text(cordsX, cordsY, text="Ahmad Almathami", fill="black", anchor=tk.CENTER, font=(current_font_family, current_font_size))

    def on_drag(event):
        global cordsX, cordsY
        cordsX = event.x
        cordsY = event.y
        coord_label.config(text=f"Coordinates: ({cordsX}, {cordsY})")
        canvas.coords(text_id, cordsX, cordsY)
    
    # Bind the drag function to the canvas
    canvas.tag_bind(text_id, "<B1-Motion>", on_drag)

    # Functions to increase or decrease the text size
    def increase_size():
        global current_font_size
        current_font_size += 2
        canvas.itemconfig(text_id, font=(current_font_family, current_font_size))
    
    def decrease_size():
        global current_font_size
        current_font_size -= 2
        canvas.itemconfig(text_id, font=(current_font_family, current_font_size))

    # Function to browse for fonts
    def browse_font():
        font_choice = tkFont.families()
        font_choice_window = tk.Toplevel(coord_window)
        font_choice_window.title("Choose Font")
        
        listbox = tk.Listbox(font_choice_window)
        listbox.pack(fill=tk.BOTH, expand=True)

        for f in font_choice:
            listbox.insert(tk.END, f)
        
        def select_font(event):
            global current_font_family
            selected_font = listbox.get(listbox.curselection())
            current_font_family = selected_font
            canvas.itemconfig(text_id, font=(current_font_family, current_font_size))
            font_choice_window.destroy()
        
        listbox.bind("<<ListboxSelect>>", select_font)

    # Add buttons for increasing and decreasing text size
    btn_increase = tk.Button(coord_window, text="Increase Text Size", command=increase_size)
    btn_increase.pack(side=tk.LEFT)
    btn_decrease = tk.Button(coord_window, text="Decrease Text Size", command=decrease_size)
    btn_decrease.pack(side=tk.LEFT)
    
    # Add button for browsing fonts
    btn_browse_font = tk.Button(coord_window, text="Browse Fonts", command=browse_font)
    btn_browse_font.pack(side=tk.LEFT)
    
    coord_window.mainloop()




# Load the Excel file
#the path to ur excel sheet file (or u can rename it to this)
# also make sure that ur excel file have each row labeled with Name, Phone, Email, to avoid errors.
# The code will automatically look for these names to access the data correctly

template_path = 'template.jpeg' #specify the template's path
cordsX = 486 # the X of the place where the name should be printed
cordsY = 342 # the Y of the place where the name should be printed
# You can use gimp (or Paint if u r on windows) to find those coordinates
# The text is centered, so the point should be centered on the space u desire to have the names on


subject = "Your certificate" # The subject of the email
body = """ 
Congrats
""" # The body of the email (Can have multiple lines)
urEmail = 'email here' # The email you are using for sending
urPass = 'password here' # The password of that email

# Certificate generation function
def generate_certificate(name, email, font_family=current_font_family, font_size=current_font_size, ):
    font_family = current_font_family # Here specify the font path
    font_size = current_font_size
    # Open the template
    img = Image.open(cert_path.get()).convert("RGB")
    draw = ImageDraw.Draw(img)

    # Set the font
    try:
        font_path = f"{font_family}.ttf"
        font = ImageFont.truetype(font_path, font_size)
    except OSError:
        font = ImageFont.load_default()

    # Calculate text position using textbbox
    bbox = draw.textbbox((0, 0), name, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    position = (cordsX - text_width // 2, cordsY - text_height // 2)

    # Add text to image
    draw.text(position, name, fill="black", font=font)

    # Save the certificate
    output_path='certificates/'

    if not os.path.exists(output_path):
        os.makedirs(output_path)
    
    certificate_path = output_path + email + '.png'
    img.save(certificate_path)
    return certificate_path




# Email sending function
def send_email(to_email, subject, body, certificate_path, from_email=urEmail, password=urPass): 
    from_email = urEmail.get()
    password = urPass.get()
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    
    attachment = open(certificate_path, 'rb')
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % certificate_path)
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





root = tk.Tk()
root.title("Certifier - by oAmadu on github")

excel_path = tk.StringVar()
tk.Label(root, text="Select Excel File:").pack()
tk.Entry(root, textvariable=excel_path).pack()
tk.Button(root, text="Browse", command=select_excel).pack()

cert_path = tk.StringVar()
tk.Label(root, text="Select Certificate File:").pack()
tk.Entry(root, textvariable=cert_path).pack()
tk.Button(root, text="Browse", command=select_template).pack()

coord_button = tk.Button(root, text="Select Coordinates", command=coordinator, state=tk.DISABLED)
coord_button.pack()


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

# Function to send a test email 
def send_test_email():
    name = 'Test Name'
    email = urEmail.get()
    phone = 'Test Number'

    certificate_path = generate_certificate(name, email, current_font_family, current_font_size)
    send_email(email, subject.get(), body.get(), certificate_path) #(email, subject, body, path)

    messagebox.showinfo("Test Email", "Test email sent!")

# Function to send bulk emails 
def send_emails():
    filepath = excel_path.get()
    if not filepath:
        messagebox.showerror("Error", "Please select an Excel file.")
        return

    try:
        df = pd.read_excel(filepath)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel file: {e}")
        return
    for _, row in df.iterrows():
        try:
            name = row['Name']
            email = row['Email']
            phone = row['Phone']
        except KeyError as e:
            messagebox.showerror("Error", f"Missing column in Excel file: \nAlso make sure the excel sheet is in the same directory.{e}")
            return
        
        if not isinstance(email, str) or email.lower() in ["no email", ""]:
            log_error(name, phone, "Invalid email")
            continue
        
        try:
            certificate_path = generate_certificate(name, email, current_font_family, current_font_size)
            send_email(email, subject.get(), body.get(), certificate_path) #(email, subject, body, path)
        except Exception as e:
            log_error(name, phone, str(e))

    messagebox.showinfo("Bulk Emails", "Bulk emails sent!")

# Sending emails
tk.Button(root, text="Send Test Email", command=send_test_email).pack()
tk.Button(root, text="Send Bulk Emails", command=send_emails).pack()

# Email context (Subject and body)
subject = tk.StringVar()
body = tk.StringVar()
tk.Label(root, text="Email Subject").pack()
subject_entry = tk.Text(root, height=1, width=50)
subject_entry.pack()

tk.Label(root, text="Email Body").pack()
body_entry = tk.Text(root, height=10, width=50)
body_entry.pack()


root.mainloop()
