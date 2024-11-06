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

fake = Faker("he_IL")


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


def get_address_under_25_chars():
    while True:
        address = fake.address()
        if len(address) <= 25:
            return address


def get_name_under_12_chars():
    while True:
        name = fake.name()
        if len(name) <= 12:
            return name


def get_city_under_10_chars():
    while True:
        city = fake.city()
        if len(city) <= 10:
            return city


pdfmetrics.registerFont(TTFont("David", "David.ttf"))

# Input and output file paths
template_pdf_path = "./ichour_tochavout.pdf"


def create_overlay_pdf(field_values, overlay_path):
    c = canvas.Canvas(overlay_path, pagesize=letter)
    # Customize coordinates and font settings as needed
    array_of_strings = [
        "Helvetica",
        "Times-Roman",
        "Courier",
        "Helvetica-Bold",
        "Times-Bold",
    ]
    c.setFont(random.choice(array_of_strings), fake.random_int(min=9, max=12))
    c.drawString(
        fake.random_int(min=70, max=120),
        320,
        field_values.get("number_plain", ""),
    )
    c.setFont(random.choice(array_of_strings), fake.random_int(min=9, max=12))
    c.drawString(
        fake.random_int(min=260, max=290),
        fake.random_int(min=320, max=330),
        field_values.get("number_ichour"),
    )
    c.setFont("David", fake.random_int(min=9, max=12))
    c.drawString(
        425,
        fake.random_int(min=320, max=330),
        reverse_hebrew_text(field_values.get("name_rechout")),
    )
    c.setFont("David", 13)
    c.drawString(250, 267, field_values.get("id"))
    c.setFont("David", fake.random_int(min=9, max=12))
    c.drawString(480, 267, reverse_hebrew_text(field_values.get("name", "")))
    c.setFont("David", 13)
    c.drawString(330, 245, reverse_hebrew_text(field_values.get("address")))
    c.drawString(240, 222, reverse_hebrew_text(field_values.get("address2", "")))
    c.drawString(65, 193, field_values.get("date1"))
    c.drawString(210, 193, field_values.get("date2", ""))
    c.drawString(433, 193, reverse_hebrew_text(field_values.get("city")))
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


def generate_pdfs():
    fake = Faker("he_IL")
    data = []
    folder_path = "files_generated2"
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    for i in range(1, 101):
        map_item = {
            "number_plain": str(fake.random_number(digits=9, fix_len=True)),
            "number_ichour": str(fake.random_number(digits=9, fix_len=True)),
            "name_rechout": random_misrad(),
            "id": str(fake.random_number(digits=9, fix_len=True)),
            "name": get_name_under_12_chars(),
            "address": get_address_under_25_chars(),
            "address2": get_address_under_25_chars(),
            "date1": fake.date(pattern="%d/%m/%Y"),
            "date2": fake.date(pattern="%d/%m/%Y"),
            "city": get_city_under_10_chars(),
        }
        data.append(map_item)
        overlay_pdf_path = f"./overlay_{str(i)}.pdf"
        filled_pdf_path = f"./files_generated2/Ichour_tochavout_{str(i)}.pdf"
        create_overlay_pdf(map_item, overlay_pdf_path)
        merge_pdfs(template_pdf_path, overlay_pdf_path, filled_pdf_path)
    # Convert the data to a DataFrame
    df = pd.DataFrame(data)

    # Save the DataFrame to an Excel file
    df.to_excel("generated_data2.xlsx", index=False, engine="openpyxl")


generate_pdfs()
