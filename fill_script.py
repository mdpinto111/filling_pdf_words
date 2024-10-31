import pdfrw
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))

# Input and output file paths
template_pdf_path = './Service_Pages_Income_tax_itc1385.pdf'
filled_pdf_path = './Filled_Service_Pages_Income_tax_itc1385.pdf'

# Define the fields and values you want to fill
field_values = {
    'שם הנישום': 'John Doe',
    'מספר תיק במס הכנסה': '1  2  3  4  5  6  7  8  9',
    'מספר תיק ניכויים': '1  2  3  4  5  6  7  8',
    'מספר טלפון': '054  5678787',
    'כתובת העסק': 'מקור חיים 36 ירושלים',
    'משרד פקיד השומה': 'משרד בע"ם',
    'משרד פקיד השומה ניכויים': 'אחיםשלי בעם',
    'שם הצד הקשור': 'יצחק לוי',
    '(TIN) מספר זיהוי לצרכי מס בחו"ל': '654654654',
    'כתובת': 'מקור חיים 54 ירושלים',
    'מספר העסקה': '1354654654',
    'תיאור העסקה': 'מפתח מתכנת ברשות המיסיסם',
    'השיטה שננקטה': 'שיטה רגילה',
    'שיעור הרווחיות': '1500 שקל',
    'סכום העסקה': '1500 שקל',
    'העסקה המדווחת היא עסקה חד פעמית': '*',
    'העסקה מסוג שירותים המוסיפים ערך נמוך': '*',
    'העסקה מסוג שירותי שיווק': '*',
    'העסקה מסוג שירותי הפצה': '*',
    'קיים דיווח חקר תנאי שוק': '*',
    'תאריך': '12/10/2024',
    'שם': 'דוד פינטו',
    'תפקיד': 'מפתח',
    'חתימה': 'דוד פינטו',
}

def reverse_hebrew_text(text):
    # Reverse the text to handle right-to-left alignment
    return text[::-1]

def create_overlay_pdf(field_values, overlay_path):
    c = canvas.Canvas(overlay_path, pagesize=letter)
    # Customize coordinates and font settings as needed
    c.setFont("Helvetica", 10)
    c.setFont("HeiseiKakuGo-W5", 10)  # Using a built-in font that supports Hebrew
    c.drawString(400, 675, field_values.get('שם הנישום', ''))
    c.drawString(230, 675, field_values.get('מספר תיק במס הכנסה', ''))
    c.drawString(134, 675, field_values.get('מספר תיק ניכויים', ''))
    c.drawString(39, 675, field_values.get('מספר טלפון', ''))
    c.drawString(350, 650, reverse_hebrew_text(field_values.get('כתובת העסק', '')))
    c.drawString(156, 650, reverse_hebrew_text(field_values.get('משרד פקיד השומה', '')))
    c.drawString(39, 650, reverse_hebrew_text(field_values.get('משרד פקיד השומה ניכויים', '')))
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
