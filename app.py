import pandas as pd
import random
import string
import os
from flask import Flask, render_template, request, send_from_directory, jsonify
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from PIL import Image
from datetime import datetime, timedelta

app = Flask(__name__)

# Folder for saving offer letters and uploads
if not os.path.exists('offer_letters'):
    os.makedirs('offer_letters')
if not os.path.exists('uploads'):
    os.makedirs('uploads')  # Ensure the uploads directory exists

# Function to generate offer letter content
def generate_offer_letter_body(designation, start_date, end_date, confirmation_deadline):
    return f"""
We are pleased to extend an offer for you to join Infiedge Software as a {designation}. 
Your internship will commence on {start_date} and conclude on {end_date}. 

At Infiedge, we strive to nurture talent by providing hands-on exposure to real-world projects. 
During your internship, you will work closely with our experienced team, 
who will guide you through various tasks and ensure a meaningful learning experience. 

You are expected to maintain a professional attitude and adhere to the guidelines provided during the Internship. 
Your commitment to excellence will contribute to both your personal growth and the ongoing success of Infiedge Software. 

Please confirm your acceptance of this offer by signing and returning a copy of this letter to us by {confirmation_deadline}. 
Upon successful completion of your internship, you will receive a certificate of completion and recognition for your contributions. 

We are excited to welcome you to the Infiedge Software family!
"""

# Function to generate a unique ID
def generate_unique_id():
    prefix = "IE-001"
    random_letter = random.choice(string.ascii_uppercase)
    random_number = random.randint(10000, 99999)
    return f"{prefix}-{random_letter}{random_number}"

# Function to wrap and print text on the PDF canvas
def wrap_text(c, text, max_width, x, y, font="Helvetica", font_size=30, line_spacing=50):
    c.setFont(font, font_size)
    words = text.split()
    lines = []
    current_line = words[0]

    for word in words[1:]:
        if c.stringWidth(current_line + " " + word, font, font_size) < max_width:
            current_line += " " + word
        else:
            lines.append(current_line)
            current_line = word
    lines.append(current_line)

    for line in lines:
        c.drawString(x, y, line)
        y -= line_spacing  # Move to the next line
        if y < 50:  # Start a new page if text reaches the bottom margin
            c.showPage()
            c.setFont(font, font_size)
            y = letter[1] - 50
    return y

# Function to create a PDF with text wrapping and specific fields on new lines
def create_pdf_with_fields(image_file, signature_file, date, unique_id, intern_name, body_text, output_pdf_file):
    # Load the image template
    image = Image.open(image_file)
    width, height = image.size

    # Create a canvas
    c = canvas.Canvas(output_pdf_file, pagesize=(width, height))
    c.drawImage(image_file, 0, 0, width, height)

    # Set the text positions
    text_x = 40
    text_y = height - 600  # Start near the top of the page

    # Print Date, Unique ID, and Dear {intern_name} on separate lines
    c.setFont("Helvetica-Bold", 30)
    c.drawString(text_x, text_y, f"Date: {date}")
    text_y -= 70
    c.setFont("Helvetica-Bold", 30)
    c.drawString(text_x, text_y, f"Unique ID: {unique_id}")
    text_y -= 60
    c.setFont("Helvetica-Bold", 30)
    c.drawString(text_x, text_y, f"Dear {intern_name},")
    text_y -= 50  # Add extra spacing before the main text

    # Wrap and print the body text
    text_y = wrap_text(c, body_text, width - 80, text_x, text_y, line_spacing=50)

    # Add "Sincerely" and "Infiedge Software" on new lines
    c.setFont("Helvetica", 25)
    text_y -= 30  # Additional spacing before "Sincerely"
    c.drawString(text_x, text_y, "Sincerely,")
    text_y -= 120
    c.drawString(text_x, text_y, "Authorized Signature")
    text_y -= 30
    c.setFont("Helvetica-Bold", 25)
    c.drawString(text_x, text_y, "Infiedge Software (CEO) ")

    # Add the signature image if provided
    if signature_file:
        signature_width = 200
        signature_height = 80
        signature_x = width - signature_width - 1200
        signature_y = 600
        c.drawImage(signature_file, signature_x, signature_y, signature_width, signature_height)

    # Save the canvas
    c.save()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate_offer_letters', methods=['POST'])
def generate_offer_letters():
    # Get the uploaded files
    excel_file = request.files['excel_file']
    image_template_file = request.files['image_template_file']
    signature_file = request.files['signature_file']

    # Save the files
    excel_file_path = os.path.join('uploads', excel_file.filename)
    image_template_path = os.path.join('uploads', image_template_file.filename)
    signature_path = os.path.join('uploads', signature_file.filename)
    
    excel_file.save(excel_file_path)
    image_template_file.save(image_template_path)
    signature_file.save(signature_path)

    # Process the Excel file and generate the offer letters
    try:
        df = pd.read_excel(excel_file_path)
        updated_data = []

        for _, row in df.iterrows():
            intern_name = row["Name"]
            designation = row["Designation"]
            start_date = row["Start Date"].strftime("%d-%m-%Y")  # Remove timestamp and format the date
            end_date = row["End Date"].strftime("%d-%m-%Y")  # Remove timestamp and format the date
            issue_date = datetime.now().strftime("%d-%B-%Y")
            confirmation_deadline = (datetime.now() + timedelta(days=7)).strftime("%d-%m-%Y")
            unique_id = generate_unique_id()

            # Generate offer letter body
            offer_letter_body = generate_offer_letter_body(designation, start_date, end_date, confirmation_deadline)

            # Generate PDF
            pdf_file_name = f"offer_letters/{intern_name.replace(' ', '_')}_Offer_Letter.pdf"
            create_pdf_with_fields(
                image_template_path,
                signature_path,
                issue_date,
                unique_id,
                intern_name,
                offer_letter_body,
                pdf_file_name
            )

            # Append updated row with unique ID to the list
            updated_row = row.copy()
            updated_row["Unique ID"] = unique_id
            updated_row["Start Date"] = start_date  # Ensure the start date is in the proper format
            updated_row["End Date"] = end_date  # Ensure the end date is in the proper format
            updated_data.append(updated_row)

        # Convert the updated list to a DataFrame
        updated_df = pd.DataFrame(updated_data)

        # Specify the fixed name for the updated Excel file (no timestamp)
        updated_excel_file = os.path.join('updated_excel', 'updated_intern_data_with_unique_ids.xlsx')

        # Save updated Excel file
        updated_df.to_excel(updated_excel_file, index=False)

        return jsonify({"message": "Offer letters generated successfully!", "excel_file": updated_excel_file})

    except Exception as e:
        return jsonify({"error": str(e)})

if __name__ == '__main__':
    app.run(debug=True)
