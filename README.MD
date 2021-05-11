# USA Dinnerware Direct - <a href="https://www.usadinnerwaredirect.com">Link</a>

As part of my responsibilities at <a href="https://www.fiestatableware.com">The Fiesta Tableware Company </a> - we needed placards to help with inventory management.  The general idea is to make basic signage for each product that would help workers locate the product and do inventory audits.  

Our store uses the Shopify platform which allows exporting all or selected collections of the products in the system in CSV format. 

CSV's are exported and I clean them manually.  First column is Product Title, second column is Product SKU, third column is Product Variant Image URL.  

Uses Py Docx to generate the document, urllib for requesting the image via url, csv to read the export from Shopify, and OS for minor changing of working directories. 


    The heading is the SKU

    The subheading is the product title 

    Images of the product, QR Code, and USADD logo 