import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import requests
from io import BytesIO
import os
import textwrap


def download_image(url):
    response = requests.get(url)
    return Image.open(BytesIO(response.content))


def add_quote_to_image(image, quote):
    # Create a copy of the image to work with
    img_with_quote = image.copy()
    draw = ImageDraw.Draw(img_with_quote)

    # Load a font (you'll need to specify a path to a font file on your system)
    try:
        font = ImageFont.truetype("arial.ttf", 36)
    except:
        font = ImageFont.load_default()

    # Wrap text to fit image width
    margin = 20
    text_width = img_with_quote.width - 2 * margin
    wrapped_text = textwrap.fill(quote, width=30)

    # Calculate text size and position
    text_bbox = draw.textbbox((0, 0), wrapped_text, font=font)
    text_width = text_bbox[2] - text_bbox[0]
    text_height = text_bbox[3] - text_bbox[1]

    # Position text at bottom of image with some padding
    x = (img_with_quote.width - text_width) // 2
    y = img_with_quote.height - text_height - margin

    # Add semi-transparent background for text
    padding = 10
    rectangle_shape = [
        (x - padding, y - padding),
        (x + text_width + padding, y + text_height + padding)
    ]
    draw.rectangle(rectangle_shape, fill=(0, 0, 0, 128))

    # Draw text
    draw.text((x, y), wrapped_text, font=font, fill=(255, 255, 255))

    return img_with_quote


def process_quotes(excel_file, output_folder):
    # Create output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    try:
        # Try reading with openpyxl engine first (for .xlsx files)
        df = pd.read_excel(excel_file, engine='openpyxl')
    except Exception as e:
        try:
            # If openpyxl fails, try xlrd engine (for .xls files)
            df = pd.read_excel(excel_file, engine='xlrd')
        except Exception as e:
            print(f"Error reading Excel file: {str(e)}")
            print("Please ensure your Excel file is either .xlsx or .xls format")
            return

    if 'quote' not in df.columns or 'image_url' not in df.columns:
        print("Error: Excel file must contain 'quote' and 'image_url' columns")
        return

    for index, row in df.iterrows():
        quote = str(row['quote'])
        image_url = str(row['image_url'])

        try:
            # Download image
            image = download_image(image_url)

            # Add quote to image
            result_image = add_quote_to_image(image, quote)

            # Save result
            output_filename = f"{quote[:30].replace(' ', '_')}.png"
            output_path = os.path.join(output_folder, output_filename)
            result_image.save(output_path)

            print(f"Generated: {output_filename}")
        except Exception as e:
            print(f"Error processing quote: {quote}")
            print(f"Error message: {str(e)}")


# Example usage
if __name__ == "__main__":
    excel_file = "anites.xlsx"  # or "quotes.xls" if you're using old Excel format
    output_folder = "generated_quotes"
    process_quotes(excel_file, output_folder)