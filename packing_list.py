import sys
import os
import tkinter
from tkinter import messagebox, simpledialog
import pandas as pd
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib import colors
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from datetime import datetime
from PyPDF2 import PdfMerger
import subprocess
import platform

# Hide main Tkinter window
root = tkinter.Tk()
root.withdraw()

# Ask for date input
date_input = simpledialog.askstring("Enter Date", "Format: DD.MM.YYYY", initialvalue=datetime.now().strftime("%d.%m.%Y"))
try:
    date_obj = datetime.strptime(date_input, "%d.%m.%Y")
    formatted_date = date_obj.strftime("%d.%m.%Y")
except:
    messagebox.showerror("Error", "Invalid date format. Use DD.MM.YYYY (e.g. 01.01.2023).")
    sys.exit()

# Register Arial font
pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

# Define styles
styles = getSampleStyleSheet()
style_normal = ParagraphStyle('NormalWrap', fontName='Arial', fontSize=9, leading=10, alignment=TA_LEFT)
style_header = ParagraphStyle('Header', fontName='Arial', fontSize=10, leading=11, alignment=TA_CENTER, textColor=colors.white, backColor=colors.HexColor("#4F81BD"))
style_title = ParagraphStyle('Title', fontName='Arial', fontSize=24, alignment=TA_CENTER, spaceAfter=20)

# Input file and resources
excel_file = 'packing_list.xlsx'

def resource_path(relative_path):
    """Get absolute path to resource (for PyInstaller compatibility)"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.abspath(relative_path)

logo_path = resource_path('logo1.png')
logo_thm = resource_path('thmlogo.png')
output_folder = 'generated_packing_lists'
os.makedirs(output_folder, exist_ok=True)

df = pd.read_excel(excel_file, dtype={'Collab Order': str})
groups = df.groupby('Customer Name')

# Table headers
headers = ['Delivery', 'Material', 'Article', 'EAN', 'Material Name', 'Qty', 'Collab Order']

# Header/footer rendering
def header_footer(canvas, doc, customer=None, address=None):
    width, height = landscape(A4)
    
    canvas.setFont("Arial", 8)
    canvas.drawString(2*cm, height - 1*cm, "Address:")
    canvas.drawString(2*cm, height - 1.5*cm, "Phone: / Fax: ")
    canvas.drawString(2*cm, height - 2*cm, "Web: ")
    canvas.drawString(2*cm, height - 2.5*cm, "EMAIL: ")
    canvas.drawString(2*cm, height - 3*cm, "Company ID:  / Activity Code: / Tax ID:")

    canvas.setLineWidth(1)
    canvas.setStrokeColor(colors.HexColor("#4F81BD"))
    canvas.line(2*cm, height - 3.3*cm, width - 2*cm, height - 3.3*cm)

    # Dispatch info (left side)
    canvas.setFont("Arial", 12)
    canvas.drawString(2*cm, height - 4*cm, "Dispatch Location: ")
    canvas.drawString(2*cm, height - 5*cm, "Dispatch Address: ")

    # Delivery info (right side)
    right_x = width - 13*cm
    block_width = 9*cm
    if customer:
        p_cust = Paragraph(f"<b>Delivery Location:</b> {customer}", ParagraphStyle('right', fontName='Arial', fontSize=12))
        p_cust.wrapOn(canvas, block_width, 1*cm)
        p_cust.drawOn(canvas, right_x, height - 4*cm)
    if address:
        p_addr = Paragraph(f"<b>Delivery Address:</b> {address}", ParagraphStyle('right', fontName='Arial', fontSize=12))
        p_addr.wrapOn(canvas, block_width, 1*cm)
        p_addr.drawOn(canvas, right_x, height - 5*cm)

    canvas.setLineWidth(0.5)
    canvas.setStrokeColor(colors.grey)
    canvas.line(2*cm, height - 5.2*cm, width - 2*cm, height - 5.2*cm)

    # Logos
    if os.path.exists(logo_path):
        canvas.drawImage(logo_path, width - 10*cm, height - 6.5*cm, width=8*cm, height=8*cm, preserveAspectRatio=True, mask='auto')
    if os.path.exists(logo_thm):
        canvas.drawImage(logo_thm, width - 19*cm, height - 6*cm, width=8*cm, height=8*cm, preserveAspectRatio=True, mask='auto')

    # Title and date
    canvas.setFont("Arial", 18)
    canvas.drawCentredString(width / 2.5, height - 6*cm, "Customer Refill")
    canvas.drawString(width / 2, height - 6*cm, f"{formatted_date}")

    # Page number
    canvas.setFont("Arial", 7)
    canvas.drawRightString(width - 2*cm, 1*cm, f"Page {doc.page}")

# Generate PDFs
for idx, (customer, group) in enumerate(groups):
    filename = f"Refill_{customer[:30].replace(' ', '_').replace('/', '_')}.pdf"
    filepath = os.path.join(output_folder, filename)

    doc = SimpleDocTemplate(filepath, pagesize=landscape(A4), rightMargin=2*cm, leftMargin=2*cm, topMargin=6*cm, bottomMargin=2*cm)

    data = [[Paragraph(h, style_header) for h in headers]]
    for _, row in group.iterrows():
        data.append([
            Paragraph(str(int(row['Delivery'])), style_normal),
            Paragraph(str(int(row['Material'])), style_normal),
            Paragraph(str(int(row['Article'])), style_normal),
            Paragraph(str(int(row['EAN'])), style_normal),
            Paragraph(str(row['Material Name']), style_normal),
            Paragraph(str(int(row['Qty'])), style_normal),
            Paragraph(str(row['Collab Order']), style_normal),
        ])

    col_widths = [2*cm, 2*cm, 2*cm, 3.5*cm, None, 1.2*cm, 4*cm]
    table = Table(data, colWidths=col_widths, repeatRows=1)

    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#4F81BD")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('FONTSIZE', (0, 1), (-1, -1), 7),
        ('ALIGN', (4, 1), (4, -1), 'RIGHT'),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
    ]))

    elements = [table]

    # Footer: Collab Orders
    unique_orders = group['Collab Order'].dropna().astype(str).unique()
    collab_paragraph = Paragraph(f"<b>Collab Orders:</b> {', '.join(unique_orders)}", style_normal)
    elements.append(Spacer(1, 12))
    elements.append(collab_paragraph)
    elements.append(Spacer(1, 36))

    # Signatures
    signature_table = Table([
        [
            Paragraph('<para align="center">___________________________<br/>Issued, Date</para>', style_normal),
            Paragraph('<para align="center">___________________________<br/>Taken By</para>', style_normal),
            Paragraph('<para align="center">___________________________<br/>Received, Date</para>', style_normal)
        ]
    ], colWidths=[6*cm]*3)

    signature_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'BOTTOM'),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    elements.append(signature_table)

    address = group['Delivery Address'].iloc[0] if not group['Delivery Address'].isnull().all() else ""
    doc.build(elements, onFirstPage=lambda c, d: header_footer(c, d, customer, address),
                        onLaterPages=lambda c, d: header_footer(c, d, customer, address))

    print(f"✅ Generated PDF for: {customer} → {filepath}")

# Merge all generated PDFs
merged_path = os.path.join(output_folder, 'All_Refills_Combined.pdf')
merger = PdfMerger()
for file in sorted(os.listdir(output_folder)):
    if file.endswith(".pdf") and file.startswith("Refill_"):
        merger.append(os.path.join(output_folder, file))
merger.write(merged_path)
merger.close()

print(f"\n📄 Combined PDF saved as: {merged_path}")

# Open output folder
def open_folder(folder):
    if platform.system() == "Windows":
        os.startfile(folder)
    elif platform.system() == "Darwin":
        subprocess.Popen(["open", folder])
    else:
        subprocess.Popen(["xdg-open", folder])

open_folder(output_folder)
