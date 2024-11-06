import pdfrw
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.ttfonts import TTFont
from faker import Faker
import random
import os
from data import misradim, jobs, shitot, transactions_description
import pandas as pd


def reverse_hebrew_text(text):
    words = text.split()
    for index, item in enumerate(words):
        if not item.isdigit():  # Check if the item is a word (all alphabetic)
            words[index] = item[::-1]  # Reverse the word

    # Reverse the list of words
    words.reverse()
    sentence = " ".join(words)

    return sentence


def random_misrad():
    random_index = random.randint(1, len(misradim) - 1) - 1
    return misradim[random_index]


def random_job():
    random_index = random.randint(1, 50) - 1
    return jobs[random_index]


def random_transaction_desc():
    random_index = random.randint(1, 100) - 1
    return transactions_description[random_index]


def random_shita():
    random_index = random.randint(1, 4) - 1
    return shitot[random_index]


# Initialize the Faker object with Hebrew locale
fake = Faker("he_IL")
fake_english = Faker("en_US")

pdfmetrics.registerFont(TTFont("David", "David.ttf"))

# Input and output file paths
template_pdf_path = "./Service_Pages_Income_tax_itc1385.pdf"


def create_overlay_pdf(field_values, overlay_path):
    c = canvas.Canvas(overlay_path, pagesize=letter)
    # Customize coordinates and font settings as needed
    c.setFont("Helvetica", 10)
    c.drawString(200, 725, field_values.get("שנה", ""))
    c.drawString(230, 675, field_values.get("מספר תיק במס הכנסה", ""))
    c.drawString(134, 675, field_values.get("מספר תיק ניכויים", ""))
    c.drawString(65, 675, field_values.get("מספר טלפון", ""))
    c.setFont("David", fake.random_int(min=9, max=13))
    c.drawString(
        fake.random_int(min=350, max=440),
        675,
        reverse_hebrew_text(field_values.get("שם הנישום")),
    )
    c.setFont("David", fake.random_int(min=9, max=13))
    c.drawString(
        fake.random_int(min=265, max=340),
        650,
        reverse_hebrew_text(field_values.get("כתובת העסק")),
    )
    c.setFont("David", 9)  # Set font to 'David' with size 12
    c.drawString(152, 650, reverse_hebrew_text(field_values.get("משרד פקיד השומה")))
    c.drawString(
        35, 650, reverse_hebrew_text(field_values.get("משרד פקיד השומה ניכויים"))
    )
    array_of_strings = [
        "Helvetica",
        "Times-Roman",
        "Courier",
        "Helvetica-Bold",
        "Times-Bold",
    ]
    c.setFont(
        random.choice(array_of_strings), fake.random_int(min=9, max=12)
    )  # Set font to 'David' with size 12
    c.drawString(430, 600, field_values.get("שם הצד הקשור"))
    c.setFont("David", 11)  # Set font to 'David' with size 12
    c.drawString(280, 600, field_values.get('(TIN) מספר זיהוי לצרכי מס בחו"ל'))
    c.drawString(40, 600, reverse_hebrew_text(field_values.get("כתובת")))
    c.setFont("David", 12)  # Set font to 'David' with size 12
    c.drawString(280, 560, reverse_hebrew_text(field_values.get("מספר העסקה")))
    c.setFont("David", 11)  # Set font to 'David' with size 12
    c.drawString(33, 540, reverse_hebrew_text(field_values.get("1תיאור העסקה")))
    c.drawString(33, 520, reverse_hebrew_text(field_values.get("2תיאור העסקה")))
    c.drawString(40, 500, reverse_hebrew_text(field_values.get("1השיטה שננקטה")))
    c.drawString(40, 480, reverse_hebrew_text(field_values.get("2השיטה שננקטה")))
    c.setFont("David", 12)  # Set font to 'David' with size 12
    c.setFont(random.choice(array_of_strings), fake.random_int(min=9, max=12))
    c.drawString(
        fake.random_int(min=100, max=300),
        460,
        reverse_hebrew_text(field_values.get("1שיעור הרווחיות")),
    )
    c.setFont(random.choice(array_of_strings), fake.random_int(min=9, max=12))
    c.drawString(
        fake.random_int(min=100, max=300),
        440,
        reverse_hebrew_text(field_values.get("2שיעור הרווחיות")),
    )
    c.setFont(random.choice(array_of_strings), fake.random_int(min=9, max=12))
    c.drawString(
        fake.random_int(min=100, max=300),
        420,
        reverse_hebrew_text(field_values.get("סכום העסקה")),
    )
    c.setFont("David", 11)  # Set font to 'David' with size 12
    c.drawString(
        130 if field_values.get("העסקה המדווחת היא עסקה חד פעמית") else 95, 380, "x"
    )
    c.drawString(
        130 if field_values.get("העסקה מסוג שירותים המוסיפים ערך נמוך") else 95,
        356,
        "x",
    )
    c.drawString(130 if field_values.get("העסקה מסוג שירותי שיווק") else 95, 334, "x")
    c.drawString(130 if field_values.get("העסקה מסוג שירותי הפצה") else 95, 313, "x")
    c.drawString(130 if field_values.get("קיים דיווח חקר תנאי שוק") else 95, 292, "x")
    c.setFont("David", 12)  # Set font to 'David' with size 12
    c.drawString(450, 190, field_values.get("תאריך", ""))
    c.drawString(330, 190, reverse_hebrew_text(field_values.get("שם")))
    c.drawString(165, 190, field_values.get("תפקיד"))
    c.drawString(70, 190, reverse_hebrew_text(field_values.get("חתימה")))
    c.save()


def merge_pdfs(template_path, overlay_path, output_path):
    template_pdf = pdfrw.PdfReader(template_path)
    overlay_pdf = pdfrw.PdfReader(overlay_path)
    for page, overlay_page in zip(template_pdf.pages, overlay_pdf.pages):
        merger = pdfrw.PageMerge(page)
        merger.add(overlay_page).render()
    pdfrw.PdfWriter(output_path, trailer=template_pdf).write()
    # Remove the overlay file
    if os.path.exists(overlay_path):
        os.remove(overlay_path)


def get_job_under_22_chars():
    while True:
        job = fake_english.job()
        if len(job) <= 22:
            return job


def get_company_under_22_chars():
    while True:
        company = fake_english.company()
        if len(company) <= 22:
            return company


def generate_pdfs():
    data = []
    folder_path = "files_generated"
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    for i in range(1, 2):
        name = fake.name()
        map_item = {
            "שנה": str(fake.year()),
            "שם הנישום": name,
            "מספר תיק במס הכנסה": "  ".join(
                str(fake.random_number(digits=9, fix_len=True))
            ),
            "מספר תיק ניכויים": "  ".join(
                str(fake.random_number(digits=8, fix_len=True))
            ),
            "מספר טלפון": fake.phone_number(),
            "כתובת העסק": fake.address(),
            "משרד פקיד השומה": random_misrad(),
            "משרד פקיד השומה ניכויים": random_misrad(),
            "שם הצד הקשור": get_company_under_22_chars(),
            '(TIN) מספר זיהוי לצרכי מס בחו"ל': " ".join(
                str(fake.random_number(digits=9, fix_len=True))
            ),
            "כתובת": fake.address(),
            "מספר העסקה": " ".join(str(fake.random_number(digits=8, fix_len=True))),
            "1תיאור העסקה": random_transaction_desc(),
            "2תיאור העסקה": random_transaction_desc(),
            "1השיטה שננקטה": random_shita(),
            "2השיטה שננקטה": random_shita(),
            "1שיעור הרווחיות": str(fake.random_int(min=10000, max=100000)),
            "2שיעור הרווחיות": str(fake.random_int(min=10000, max=100000)),
            "סכום העסקה": str(fake.random_int(min=50000, max=200000)),
            "העסקה המדווחת היא עסקה חד פעמית": random.choice([True, False]),
            "העסקה מסוג שירותים המוסיפים ערך נמוך": random.choice([True, False]),
            "העסקה מסוג שירותי שיווק": random.choice([True, False]),
            "העסקה מסוג שירותי הפצה": random.choice([True, False]),
            "קיים דיווח חקר תנאי שוק": random.choice([True, False]),
            "תאריך": fake.date(pattern="%d/%m/%Y"),
            "שם": name,
            "תפקיד": get_job_under_22_chars(),
            "חתימה": name,
        }
        data.append(map_item)
        overlay_pdf_path = f"./overlay_{str(i)}.pdf"
        filled_pdf_path = f"./files_generated/Filled_Income_tax_itc1385_{str(i)}.pdf"
        create_overlay_pdf(map_item, overlay_pdf_path)
        merge_pdfs(template_pdf_path, overlay_pdf_path, filled_pdf_path)
    # Convert the data to a DataFrame
    df = pd.DataFrame(data)

    # Save the DataFrame to an Excel file
    df.to_excel("generated_data.xlsx", index=False, engine="openpyxl")


generate_pdfs()
