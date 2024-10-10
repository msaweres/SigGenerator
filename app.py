import base64

from flask import Flask,render_template,request,redirect,send_file,url_for
from docx import Document
import difflib  # For comparing strings and finding close matches
import io
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import base64  # For encoding the file data as a base64 string
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition



app = Flask(__name__)

# Define a dictionary to map programs to addresses
PROGRAM_ADDRESS_MAP = {
    "AMBER HOUSE": ("516 31st St","Oakland, CA 94609"),
    "BRIDGE/OPPORTUNITY VILLAGE": ("515 E 18th St.","Antioch, CA 94509"),
    "DELTA LANDING": ("2101 Loveridge Rd","Pittsburg, CA 94565"),
    "DON BROWN SHELTER": ("1401 West 4th Street","Antioch, CA 94509"),
    "FAIRFIELD TRANSITIONAL HOUSING": ("345 E Travis Blvd (Unit A)","Fairfield, CA 94533"),
    "FREMONT NAVIGATION CENTER": ("3300 Capitol Ave, Building C","Fremont, CA 94538"),
    "FREMONT WELLNESS CENTER": ("40963 Grimmer Blvd","Fremont, CA 94538"),
    "GOLDEN BEAR HOTEL": ("1620 San Pablo Ave","Berkeley, CA 94702"),
    "HAYWARD NAVIGATION CENTER": ("3788 Depot Road","Hayward, CA 94545"),
    "HEDCO": ("590 B Street","Hayward, CA 94541"),
    "HENRY ROBINSON HOTEL": ("559 16th Street","Oakland, CA 94612"),
    "HENRY ROBINSON": ("559 16th Street","Oakland, CA 94612"),
    "HYPE": ("238 Capitol Street","Salinas, CA 93901"),
    "MARK TWAIN SENIOR COMMUNITY": ("3525 Lyon Avenue","Oakland, CA 94601"),
    "NEVIN HOUSE": ("3221 Nevin Avenue","Richmond, CA 94804"),
    "NORTH COUNTY HRC": ("2809 Telegraph Ave","Berkeley, CA 94705"),
    "PERALTA": ("920 Peralta Street","Oakland, CA 94607"),
    "PIEDMONT PLACE": ("55 MacArthur Blvd","Oakland, CA 94610"),
    "ROSEWOOD": ("508 Alabama Street","Vallejo, CA 94590"),
    "CEDAR": ("4600 47th Avenue, Suite # 111 & 211","Sacramento, CA 95824"),
    "SYCAMORE - 9343": ("9343 Tech Center Drive, Suite # 175 & 185","Sacramento, CA 95826"),
    "SYCAMORE": ("9333 Tech Center Dr, Ste 100","Sacramento, CA 95826"),
    "WILLOW": ("7171 Bowling Drive, Suite 300","Sacramento, CA 95823"),
    "STAIR": ("650 Cedar Street","Berkeley, CA 94710"),
    "ST. REGIS": ("23950 Mission Blvd","Hayward, CA 94544"),
    "ST REGIS": ("23950 Mission Blvd","Hayward, CA 94544"),
    "THE HOLLAND": ("641 West Grand Avenue","Oakland, CA 94612"),
    "THUNDER ROAD": ("390 40th Street","Oakland, CA 94609"),
    "TOWNE HOUSE": ("629 Oakland Ave","Oakland, CA 94611"),
    "VALLEY WELLNESS CENTER": ("3900 Valley Ave","Pleasanton, CA 94566"),
    "WOODROE": ("22505 Woodroe Ave","Hayward, CA 94541")
}

@app.route("/reset")
def reset():
    global parsed_data
    parsed_data = {}  # Clear the parsed data
    return redirect(url_for("index"))

# Comment out the email-sending function or remove it entirely if not needed
# def send_email_with_attachment(to_email, subject, body, attachment, filename):
#     try:
#         # Create the email message
#         msg = MIMEMultipart()
#         msg['From'] = smtp_user
#         msg['To'] = to_email
#         msg['Subject'] = subject
#
#         # Attach the email body
#         msg.attach(MIMEText(body, 'plain'))
#
#         # Attach the file
#         part = MIMEBase('application', 'octet-stream')
#         part.set_payload(attachment.getvalue())  # Use in-memory file
#         encoders.encode_base64(part)
#         part.add_header('Content-Disposition', f'attachment; filename={filename}')
#         msg.attach(part)
#
#         # Connect to the SMTP server
#         server = smtplib.SMTP(smtp_server, smtp_port)
#         server.starttls()
#         server.login(smtp_user, smtp_password)
#
#         # Send the email
#         server.send_message(msg)
#         server.quit()
#         print(f"Email sent to {to_email} successfully.")
#         return True
#     except Exception as e:
#         print(f"Failed to send email to {to_email}. Error: {e}")
#         return False

# Get the current script directory and set the template path
script_dir = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_FILE_PATH = os.path.join(script_dir,"Email Signatures TEMPLATE.docx")

# Global variable to store parsed data (similar to original Tkinter app)
parsed_data = {}


def get_program_address ( program_name ):
    normalized_program_name = program_name.strip().upper()
    if normalized_program_name in PROGRAM_ADDRESS_MAP:
        return PROGRAM_ADDRESS_MAP[normalized_program_name]
    closest_matches = difflib.get_close_matches(normalized_program_name,PROGRAM_ADDRESS_MAP.keys(),n=3,cutoff=0.6)
    return None


def parse_row_data(row_data):
    # Print the raw row data for debugging purposes
    print(f"Raw row data: {row_data}")

    # Split the row data using tab separator
    fields = row_data.split("\t")
    print(f"Split fields: {fields}")

    # Check if the expected number of fields is present
    if len(fields) < 14:
        print(f"Error: Expected at least 14 fields but got {len(fields)}. Fields: {fields}")
        return None

    # Extract the required fields
    name = fields[0].title().replace(" ", "")  # Capitalize and remove spaces for the filename
    program = fields[2].upper().strip()  # Program should be in uppercase and trimmed
    position = fields[4]
    phone_number = fields[10]

    # Check if the program has a corresponding address, using the improved function
    address_tuple = get_program_address(program)
    if not address_tuple:
        print(f"Error: Program '{program}' not found in PROGRAM_ADDRESS_MAP.")
        return None  # Exit if no matching address is found

    # Split the program address into street address and city/state/zip
    street_address, city_state_zip = address_tuple

    # Return a dictionary of parsed data
    parsed_data = {
        "Name": fields[0].title(),  # Capitalized Name for Filename
        "{NAME}": fields[0].title(),  # Full Name for Placeholder Replacement
        "{POSITION}": position,
        "{PHONE NUMBER}": phone_number,
        "{PROGRAM}": program,
        "{ADDRESS}": street_address,  # Street address goes into {ADDRESS}
        "{ADDRESS2}": city_state_zip  # City/State/Zip goes into {ADDRESS2}
    }

    # Print the parsed data for debugging
    print(f"Parsed Data: {parsed_data}")

    return parsed_data



@app.route("/",methods=["GET","POST"])
def index ():
    global parsed_data
    if request.method == "POST":
        row_data = request.form.get("row_data","").strip()
        if row_data:
            # Parse the row data
            parsed_data = parse_row_data(row_data)
            if not parsed_data:
                return render_template("index.html",error="Failed to parse the row data. Check the input format.")

            # Redirect to the edit page with the parsed data
            return redirect(url_for("edit_fields"))

    return render_template("index.html")


# Remove or comment out the email-sending logic in the `edit_fields` route
@app.route("/edit",methods=["GET","POST"])
def edit_fields ():
    global parsed_data
    if request.method == "POST":
        # Update parsed_data with the values from the form
        for key in parsed_data.keys():
            if key in request.form:
                parsed_data[key] = request.form[key]

        # Generate the filename dynamically using the "Name" field
        name_for_filename = parsed_data.get("Name","Unknown").replace(" ","")
        dynamic_filename = f"{name_for_filename}Signature.docx"

        # Fill the template with the updated values and get the in-memory file
        filled_document = fill_template_with_data(parsed_data,TEMPLATE_FILE_PATH)

        if filled_document:
            # Comment out the email sending code to disable auto-emailing
            # to_email = parsed_data.get("Email")
            # if to_email:
            #     subject = "Your New Email Signature"
            #     body = f"Dear {parsed_data['Name']},\n\nPlease find attached your new email signature."
            #
            #     # SMTP configuration (Update with your SMTP server details)
            #     smtp_server = "smtp.gmail.com"
            #     smtp_port = 587
            #     smtp_user = "your-email@gmail.com"
            #     smtp_password = "your-email-password"
            #
            #     # Send the email
            #     send_email_with_attachment(to_email, subject, body, filled_document, dynamic_filename, smtp_server, smtp_port, smtp_user, smtp_password)

            # Serve the filled document as a downloadable file with the dynamic filename
            return send_file(filled_document,as_attachment=True,download_name=dynamic_filename)

    return render_template("edit_fields.html",data=parsed_data)


def fill_template_with_data ( parsed_data,template_file_path ):
    document = Document(template_file_path)
    for paragraph in document.paragraphs:
        for key,value in parsed_data.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key,value)

    filled_file = io.BytesIO()
    document.save(filled_file)
    filled_file.seek(0)
    return filled_file


if __name__ == "__main__":
    app.run(debug=True)
