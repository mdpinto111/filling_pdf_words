import pdfrw
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
import re
from reportlab.pdfbase.ttfonts import TTFont

pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
pdfmetrics.registerFont(TTFont('David', 'David.ttf'))

# Input and output file paths
template_pdf_path = './Service_Pages_Income_tax_itc1385.pdf'
filled_pdf_path = './Filled_Service_Pages_Income_tax_itc1385.pdf'

# Define the fields and values you want to fill
field_values = {
    'שנה': '2024',
    'שם הנישום': 'משה יצחקי',
    'מספר תיק במס הכנסה': '1  2  3  4  5  6  7  8  9',
    'מספר תיק ניכויים': '1  2  3  4  5  6  7  8',
    'מספר טלפון': '054  5678787',
    'כתובת העסק': 'מקור חיים, ירושלים ת.ד 512',
    'משרד פקיד השומה': 'משרד בע"ם',
    'משרד פקיד השומה ניכויים': 'אחים שלי בעם',
    'שם הצד הקשור': 'לוי יצחק בע"מ',
    '(TIN) מספר זיהוי לצרכי מס בחו"ל': '654654654',
    'כתובת': 'מקור חיים 54 ירושלים',
    'מספר העסקה': '1354654654',
    '1תיאור העסקה': 'מפתח מתכנת ברשות המיסים',
    '2תיאור העסקה': 'מפתח מתכנת ברשות המיסים',
    '1השיטה שננקטה': 'שיטה רגילה',
    '2השיטה שננקטה': 'שיטה רגילה',
    '1שיעור הרווחיות': '80000',
    '2שיעור הרווחיות': '80000',
    'סכום העסקה': '80000',
    'העסקה המדווחת היא עסקה חד פעמית': '*',
    'העסקה מסוג שירותים המוסיפים ערך נמוך': '*',
    'העסקה מסוג שירותי שיווק': '*',
    'העסקה מסוג שירותי הפצה': '*',
    'קיים דיווח חקר תנאי שוק': '*',
    'תאריך': '12/10/2024',
    'שם': 'דוד יצחק',
    'תפקיד': 'מפתח',
    'חתימה': 'דוד יצחק',
}

def reverse_hebrew_text(text):
    words = text.split()
    for index, item in enumerate(words):
        if not item.isdigit():  # Check if the item is a word (all alphabetic)
            words[index] = item[::-1]  # Reverse the word

    # Reverse the list of words
    words.reverse()
    sentence = " ".join(words)
    
    # Join the words back together with spaces, preserving the original word order
    return sentence

def create_overlay_pdf(field_values, overlay_path):
    c = canvas.Canvas(overlay_path, pagesize=letter)
    # Customize coordinates and font settings as needed
    c.setFont("Helvetica", 10)
    c.drawString(200, 725, field_values.get('שנה', ''))
    c.drawString(230, 675, field_values.get('מספר תיק במס הכנסה', ''))
    c.drawString(134, 675, field_values.get('מספר תיק ניכויים', ''))
    c.drawString(39, 675, field_values.get('מספר טלפון', ''))
    c.setFont("David", 12)  # Set font to 'David' with size 12
    c.drawString(400, 675, reverse_hebrew_text(field_values.get("שם הנישום")))
    c.drawString(350, 650, reverse_hebrew_text(field_values.get("כתובת העסק")))
    c.drawString(160, 650, reverse_hebrew_text(field_values.get("משרד פקיד השומה")))
    c.drawString(50, 650, reverse_hebrew_text(field_values.get("משרד פקיד השומה ניכויים")))
    c.drawString(400, 600, reverse_hebrew_text(field_values.get("שם הצד הקשור")))
    c.drawString(280, 600, field_values.get('(TIN) מספר זיהוי לצרכי מס בחו"ל'))
    c.drawString(80, 600, reverse_hebrew_text(field_values.get('כתובת')))
    c.drawString(280, 560, reverse_hebrew_text(field_values.get('מספר העסקה')))
    c.drawString(280, 540, reverse_hebrew_text(field_values.get('1תיאור העסקה')))
    c.drawString(280, 520, reverse_hebrew_text(field_values.get('2תיאור העסקה')))
    c.drawString(280, 500, reverse_hebrew_text(field_values.get('1השיטה שננקטה')))
    c.drawString(280, 480, reverse_hebrew_text(field_values.get('2השיטה שננקטה')))
    c.drawString(280, 460, reverse_hebrew_text(field_values.get('1שיעור הרווחיות')))
    c.drawString(280, 440, reverse_hebrew_text(field_values.get('2שיעור הרווחיות')))
    c.drawString(280, 420, reverse_hebrew_text(field_values.get('סכום העסקה')))
    c.drawString(129, 380, 'x')
    c.drawString(129, 355, 'x')
    c.drawString(129, 332, 'x')
    c.drawString(129, 311, 'x')
    c.drawString(129, 290, 'x')
    c.drawString(450, 190, field_values.get('תאריך', ''))
    c.drawString(330, 190, reverse_hebrew_text(field_values.get('שם')))
    c.drawString(210, 190, reverse_hebrew_text(field_values.get('תפקיד')))
    c.drawString(70, 190, reverse_hebrew_text(field_values.get('חתימה')))
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
#print('1500 שקל')