import base64

import qrcode
import openpyxl
from io import BytesIO
from PIL import Image


# Function to create a QR code with embedded photo from the given data and save it as an image
def create_qr_code_with_photo(data, photo_data, filename):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")

    # Convert the base64-encoded image data back to an image and resize
    if photo_data:
        image = Image.open(BytesIO(photo_data))
        image = image.resize((100, 100))  # Adjust the size as needed

        # Paste the image in the QR code
        img.paste(image, (150, 150))  # Adjust the coordinates as needed

    img.save(filename)


# Load the Excel spreadsheet
excel_file = "data.xlsx"
workbook = openpyxl.load_workbook(excel_file)
sheet = workbook.active

# Assuming the data starts from row 2 (excluding headers)
start_row = 2

# Iterate through rows and create QR codes with embedded photos
for row in sheet.iter_rows(min_row=start_row):
    student_data = {
        'Serial Code': str(row[0].value) if row[0].value else '',
        'Matric No': str(row[1].value) if row[1].value else '',
        'First Name': str(row[2].value) if row[2].value else '',
        'Middle Name': str(row[3].value) if row[3].value else '',
        'Surname': str(row[4].value) if row[4].value else '',
        'Class of Degree': str(row[5].value) if row[5].value else '',
        'DOB': str(row[6].value) if row[6].value else '',
        'Date of Graduation': str(row[7].value) if row[7].value else '',
        'Course of Study': str(row[8].value) if row[8].value else '',
        'Link': "https://www.school.edu.ng/",
    }

    # Get the embedded image as bytes and encode it in base64
    photo_data = row[9].value
    photo_data_base64 = None
    if photo_data:
        photo_bytes = photo_data.read()
        photo_data_base64 = base64.b64encode(photo_bytes).decode()

    # Save the QR code as an image
    matric_no = student_data['Matric No']
    qr_code_filename = f"qr_codes/{matric_no}.png"  # You can change the path and format
    create_qr_code_with_photo(str(student_data), photo_data_base64, qr_code_filename)

print("QR codes with embedded photos generated successfully.")
