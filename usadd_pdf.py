import os, csv, urllib.request, qrcode
from PIL import Image
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import LETTER

folder = os.getcwd()

# Function to concatenate the images
# https://note.nkmk.me/en/python-pillow-concat-images/#concatenate-multiple-images-at-once
def get_concat_h_multi_resize(im_list, resample=Image.BICUBIC):
            min_height = min(im.height for im in im_list)
            im_list_resize = [im.resize((int(im.width * min_height / im.height), min_height),resample=resample)
                            for im in im_list]
            total_width = sum(im.width for im in im_list_resize)
            dst = Image.new('RGB', (total_width, min_height))
            pos_x = 0
            for im in im_list_resize:
                dst.paste(im, (pos_x, 0))
                pos_x += im.width
            return dst

with open('products_export.csv', 'r') as csv_file:
    csv_reader = csv.reader(csv_file)
    i=0
    next(csv_reader)

    pageWidth = 11 * inch
    pageHeight = 8.5 * inch

    for line in csv_reader:
        i = i + 1
        
        title = line[0]
        sku = line[1]
        # Shopify sometimes adds a ' before SKUs, this ignores if exists
        if "'" in sku:
            sku = sku[1:]
            print("' Detected, skipping first char in {sku}.  Document {page} complete. ".format(sku = sku, page = i))
        else:
            print("{sku} ready, document {page} complete. ".format(sku = sku, page = i))
            pass

        url = line[2]
        fn = "output/{0}.pdf".format(sku)

        canvas = Canvas(fn)

        # Distances are W x H
        # More info at: 
        # https://realpython.com/creating-modifying-pdf/#installing-reportlab

        canvas = Canvas(fn, pagesize=(pageWidth, pageHeight))
        canvas.setFont("Times-Bold", 85)
        canvas.drawCentredString(pageWidth/2, pageHeight-125, sku)
        canvas.setFont("Times-Roman", 26)
        canvas.drawCentredString(pageWidth/2, pageHeight-170, title)

        os.chdir(folder+"/images")
        r = urllib.request.urlopen(url)
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
        # Populate image variables
        im1 = Image.open(folder + "/images/" + sku + suffix)
        im2 = Image.open(folder + "/images/qr-codes/" + sku + "-qrcode.png")
        # Concatenate the images and save to a folder
        get_concat_h_multi_resize([im1, im2]).save(folder + "/images/concat/" + sku + '-resize.jpg')
        line_x_start = 50
        line_x_end = pageWidth - 100
        line_y = 100
        imgHeight = 4.5 * inch
        canvas.drawImage(folder+"/images/concat/"+sku+"-resize.jpg", line_x_start, 100, line_x_end, imgHeight, None, True)

        canvas.save()
