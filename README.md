# PDF Filler with Synthetic Data

This script generates and fills PDF forms with synthetic data using the `Faker` library and `ReportLab`. It also exports the generated data into an Excel file.

## Overview

The script performs the following tasks:

1. Generates synthetic data in Hebrew and English using `Faker`.
2. Fills a PDF template (`Service_Pages_Income_tax_itc1385.pdf`) with the generated data.
3. Creates overlay PDFs and merges them with the template to produce completed forms.
4. Exports the synthetic data into an Excel file named `generated_data.xlsx`.

## Requirements

The script uses several Python libraries, which are listed in the `requirements.txt` file. You can install them using the following command:

## How to Use

Clone the Repository:
git clone https://github.com/your-repo/pdf-filler-synthetic-data.git
cd pdf-filler-synthetic-data

Install Dependencies:
pip install -r requirements.txt

Run the Script:
python fill_script.py

Output:
Filled PDF forms will be saved in the files_generated folder.
An Excel file generated_data.xlsx will be created with all the generated data.

License
This project is licensed under the MIT License. See the LICENSE file for more details.
