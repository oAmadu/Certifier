# Certifier

Hi, welcome to Certifier! This project I have been working on to send certificates to the attendees of the buisness camp I volunteered in.
Wait, you're not following me on LinkedIn? Here's my [LinkedIn profile](https://www.linkedin.com/feed/update/urn:li:activity:7199866491486298112/).

Certifier is a Python script designed to generate personalized certificates from a template and send them via email. This project reads recipient details from an Excel file, customizes certificates with individual names, and emails the generated certificates automatically.

## Requirements

- Python 3.x
- pandas
- Pillow
- smtplib
- openpyxl
- arabic-reshaper
- python-bidi
- unidecode

Install the required libraries using:
```bash
pip install pandas pillow openpyxl arabic-reshaper python-bidi unidecode
```
## Basic Tutorial

1. **Prepare the Excel File**: Create an Excel file (`sheet.xlsx`) with columns for `Name`, `Email`, and `Phone`.

2. **Customize the Script**: Modify the `template_path`, `cordsX`, `cordsY`, `certFont`, `fontSize`, `subject`, and `body` variables in the script to fit your needs.

3. **Run the Script**: Execute the script to generate certificates and send them via email:
```bash
python certifier.py
```
## Future Plans

- **GUI Integration**: Add a graphical user interface to make the script more user-friendly.
- **Error Handling and Logging**: Improve error handling and logging for better troubleshooting.

## Contributing

Feel free to contribute to this project by opening issues or submitting pull requests. Your contributions are welcome!

## License

This project is licensed under the MIT License.
