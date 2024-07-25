# Certifier

Hi, welcome to Certifier! This project I have been working on to send certificates to the attendants of the workshop I was a volunteer in.
Wait, you're not following me on LinkedIn? Here's my [LinkedIn profile](https://www.linkedin.com/feed/update/urn:li:activity:7199866491486298112/).

Certifier is a Python script designed to generate personalized certificates from a template and send them via email. 
Now with a GUI for easier coordination and customization!

## Features

- Generate personalized certificates
- Send certificates via email
- GUI for selecting template and Excel files
- Coordinate text placement on the certificate
- Customize email subject and body

## Requirements

- Python 3.x
- pandas
- Pillow
- smtplib
- openpyxl
- arabic-reshaper
- python-bidi
- tkinter

Install the required libraries using:
```bash
pip install pandas pillow openpyxl arabic-reshaper python-bidi tkinter
```
## Installation

3. **Clone the repository:**
```bash
git clone https://github.com/oAmadu/Certifier.git
cd Certifier
```
3. **Install the required libraries:** `pip install pandas pillow openpyxl arabic-reshaper python-bidi tkinter`

3. **Run the script:** `python certifier.py`

## Basic Tutorial



1. **Select Excel File:**: Browse and select the Excel file with columns for `Name`, `Email`, and `Phone`.

2. **Select Template File**: Browse and select the certificate template image.

3. **Coordinate Text Placement**: Use the GUI to place the text on the template and adjust its size and font.

4. **Enter Email Credentials**: Input your email and generated app password.

5. **Customize Email Subject and Body**: Enter the subject and body of the email.

6. **Send Emails**: Use the GUI buttons to send a test email or bulk emails.

## Future Plans

- **GUI Enhancing**: Improve the GUI with more customization options
- **Error Handling and Logging**: Enhance error handling and logging

## Contributing

Feel free to contribute to this project by opening issues or submitting pull requests. Your contributions are welcome!

## License

This project is licensed under the MIT License.

## Contact Information

For support or collaboration, feel free to reach out on [LinkedIn](https://www.linkedin.com/in/aymt7mi/).

---

Thank you for using Certifier v1!
