import docx, os, csv, urllib.request, qrcode
from docx.shared import Inches, Pt, RGBColor
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT

# Fetch the current working directory of the script
folder = os.getcwd()

# Open the CSV file and read it
with open('products_export.csv', 'r') as csv_file:
    csv_reader = csv.reader(csv_file)
    
    # Skips row 1 (the column titles)
    next(csv_reader)

    # Read the CSV row by row
    for line in csv_reader:

        # Create a new document
        doc = docx.Document()
        section = doc.sections[-1]

        # Rotate the document to landscape mode
        new_width, new_height = section.page_height, section.page_width
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = new_width
        section.page_height = new_height

        # Read the CSV column by column
        title = line[0]
        sku = line[1]
        url = line[2]

        # Fetch the image from the URL
        r = urllib.request.urlopen(url)

        # Change directory to images folder
        os.chdir(folder + "/images")

        # Check the ending on the image
        if ".png?" in url:
            suffix = ".png"
        elif ".jpg?" in url:
            suffix = ".jpg"
        else:
            suffix = ".jpeg"

        # Save the image
        with open(sku + suffix, "wb") as f:
            f.write(r.read())

        # Change directory to QR Code folder
        os.chdir(folder + "/images/qr-codes") 

        # Generate the QR Code
        img = qrcode.make(sku)
        img.save(sku + "-qrcode.png")

        # Change directory to original location
        os.chdir(folder)

        # SKU line (the sku of the product) & styling
        sku_head = doc.add_heading(sku, 0)
        sku_head.style = doc.styles.add_style('Style Name SKU', WD_STYLE_TYPE.PARAGRAPH)
        sku_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
        font = sku_head.style.font
        font.name = 'Times New Roman'
        font.size = Pt(65)
        font.bold = True
        font.italic = False

        # Title line (the name of the product) & styling
        title_head = doc.add_heading(title, 1)
        title_head.style = doc.styles.add_style('Style Name TITLE', WD_STYLE_TYPE.PARAGRAPH)
        title_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
        font = title_head.style.font
        font.name = 'Calibri'
        font.size = Pt(28)
        font.color.rgb = RGBColor(0,0,0)
        font.bold = False

        # Add three images
        doc.add_picture("images/" + sku + suffix, width=Inches(3.5), height=Inches(3.5))
        doc.add_picture("images/qr-codes/" + sku + "-qrcode.png", width=Inches(2.5))
        doc.add_picture("usadd-logo.png", width=Inches(2.5))

        # Save the document to the output folder
        doc.save(folder + "/output/" + sku + ".docx")
