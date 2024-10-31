import pdfrw
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

# Input and output file paths
template_pdf_path = './Service_Pages_Income_tax_itc1385.pdf'
filled_pdf_path = './Filled_Service_Pages_Income_tax_itc1385.pdf'

# Define the fields and values you want to fill
field_values = {
    'שם הנישום': 'John Doe',
    'מספר תיק במס הכנסה': '1  2  3  4  5  6  7  8  9',
    'מספר תיק ניכויים': '1  2  3  4  5  6  7  8',
    'מספר טלפון': '054  5678787',
    # Add more fields and values as needed
}

def create_overlay_pdf(field_values, overlay_path):
    c = canvas.Canvas(overlay_path, pagesize=letter)
    # Customize coordinates and font settings as needed
    c.setFont("Helvetica", 10)
    c.drawString(400, 675, field_values.get('שם הנישום', ''))
    c.drawString(230, 675, field_values.get('מספר תיק במס הכנסה', ''))
    c.drawString(134, 675, field_values.get('מספר תיק ניכויים', ''))
    c.drawString(39, 675, field_values.get('מספר טלפון', ''))
    c.save()

def merge_pdfs(template_path, overlay_path, output_path):
    template_pdf = pdfrw.PdfReader(template_path)
    overlay_pdf = pdfrw.PdfReader(overlay_path)
    for page, overlay_page in zip(template_pdf.pages, overlay_pdf.pages):
        merger = pdfrw.PageMerge(page)
        merger.add(overlay_page).render()
    pdfrw.PdfWriter(output_path, trailer=template_pdf).write()

# Create overlay PDF
overlay_pdf_path = './overlay.pdf'
create_overlay_pdf(field_values, overlay_pdf_path)

# Merge and output the filled PDF
merge_pdfs(template_pdf_path, overlay_pdf_path, filled_pdf_path)

print(f"Filled PDF saved to: {filled_pdf_path}")
