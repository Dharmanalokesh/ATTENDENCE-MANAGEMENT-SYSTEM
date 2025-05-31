import qrcode
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import os

# Step 1: Create the output folder if it doesn't exist
output_folder = "qrcode3"
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Step 2: Read the Excel file
try:
    df = pd.read_excel("pin2.xlsx")
except FileNotFoundError:
    print("Error: The file 'pin2.xlsx' was not found. Please ensure it is in the same directory as this script.")
    exit()

# Step 3: Check if the Excel file has data
if df.empty:
    print("Error: The Excel file 'pin2.xlsx' is empty.")
    exit()

# Step 4: Verify required columns exist
required_columns = ["PIN (Roll.No)", "NAME", "BRANCH", "COURSE"]
missing_columns = [col for col in required_columns if col not in df.columns]
if missing_columns:
    print(f"Error: The following required columns were not found in the Excel file: {missing_columns}")
    print(f"Available columns: {df.columns}")
    exit()

# Step 5: Loop through each row in the Excel file
for index, row in df.iterrows():
    # Extract the branch and roll number
    branch = str(row["BRANCH"]).strip()
    roll_no = str(row["PIN (Roll.No)"]).strip()

    # Create the data string in the format BRANCH_PIN (Roll.No)_3-1
    data = f"{branch}_{roll_no}_3-1"

    # Step 6: Generate the QR code
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white")

    # Convert qr_img to RGB mode to match the background
    qr_img = qr_img.convert("RGB")

    # Step 7: Create a custom background with yellow and blue sections
    qr_size = qr_img.size[0]  # 340 pixels (based on box_size and border)
    bg_width = qr_size + 20  # Add padding
    bg_height = qr_size + 60  # Extra height for the text
    background = Image.new("RGB", (bg_width, bg_height), "white")

    # Draw the yellow and blue sections
    draw = ImageDraw.Draw(background)
    yellow_height = qr_size + 20
    draw.rectangle([0, 0, bg_width, yellow_height], fill="#FFC107")  # Yellow
    draw.rectangle([0, yellow_height, bg_width, bg_height], fill="#3F51B5")  # Blue

    # Step 8: Paste the QR code onto the background
    qr_position = ((bg_width - qr_size) // 2, 10)  # Center with 10px top margin
    background.paste(qr_img, qr_position)

    # Step 9: Add the text at the bottom
    try:
        font = ImageFont.truetype("arial.ttf", 20)
    except:
        font = ImageFont.load_default()

    # Format the text as "BRANCH - PIN (Roll.No) - 3-1" for display
    display_text = f"{branch} - {roll_no} - 3-1"
    text_bbox = draw.textbbox((0, 0), display_text, font=font)
    text_width = text_bbox[2] - text_bbox[0]
    text_height = text_bbox[3] - text_bbox[1]

    # Center the text in the blue section
    text_x = (bg_width - text_width) // 2
    text_y = yellow_height + (bg_height - yellow_height - text_height) // 2
    draw.text((text_x, text_y), display_text, fill="white", font=font)

    # Step 10: Save the QR code image in the qrcode3 folder
    # Use a sanitized filename (replace special characters if needed)
    safe_filename = data.replace("_", "-")  # Replace underscores with hyphens for the filename
    output_path = os.path.join(output_folder, f"qr_code_{safe_filename}.png")
    background.save(output_path)
    print(f"Generated QR code for {data} and saved as {output_path}")

print("All QR codes have been generated and saved in the 'qrcode3' folder.")