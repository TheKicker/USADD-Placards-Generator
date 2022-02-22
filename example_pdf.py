from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.units import inch, cm
from reportlab.lib.pagesizes import LETTER

canvas = Canvas("hello.pdf")
canvas = Canvas("hello.pdf", pagesize=(8.5 * inch, 11 * inch))
canvas = Canvas("hello.pdf", pagesize=LETTER)
canvas.setFont("Courier", 36)
canvas.drawString(72, 72, "Hello, World")
canvas.setFont("Times-Roman", 18)
canvas.drawString(1 * inch, 10 * inch, "Times New Roman (18 pt)")
canvas.save()