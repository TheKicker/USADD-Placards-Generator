import docx, os, csv, urllib.request, qrcode
from docx.shared import Inches

folder = os.getcwd()

with open('products_export.csv', 'r') as csv_file:
    csv_reader = csv.reader(csv_file)
    
    # Skips row 1 (the column titles)
    next(csv_reader)

    # Read the CSV row by row
    for line in csv_reader:

        # Open the Placard Template
        doc = docx.Document('placard.docx')
        # Read the heading data
        skuHEADING = doc.paragraphs[0]
        titleHEADING = doc.paragraphs[1]

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

        # Replace the heading with SKU, subheading with title, add images
        skuHEADING.text = sku
        titleHEADING.text = title
        doc.add_picture("images/" + sku + suffix, width=Inches(3.5), height=Inches(3.5))
        doc.add_picture("images/qr-codes/" + sku + "-qrcode.png", width=Inches(2.5))
        doc.add_picture("usadd-logo.png", width=Inches(2.5))

        # Save the document to the output folder
        doc.save(folder + "/output/" + sku + ".docx")
