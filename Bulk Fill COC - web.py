# """
# PDF Form Filler - Streamlit Web UI

# ==============================
# 📌 Required Python Modules
# ==============================
# pip install streamlit
# pip install pandas
# pip install PyMuPDF
# pip install openpyxl   # Needed for reading Excel (.xlsx)
# streamlit run "C:\Users\fchen178\OneDrive - Vialto\Fanglin\Techhhhhh\pdf generator\CoC\Bulk Fill COC - web.py"

# """

import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import os
import tempfile
import zipfile
import glob
from datetime import date, datetime
import pikepdf
from pikepdf import Pdf
from fitz import TextWriter

# ===============================
# Function to fill and flatten the PDF
# ===============================
def fill_and_flatten_pdf(template_pdf, output_pdf, data):
    from datetime import datetime, date


    # --- Gregorian → Japanese conversion (uses parse_date) ---
    def parse_date(date_value):
        """Parse different date formats into a datetime.date."""
        if isinstance(date_value, str):
            for fmt in ("%y/%m/%d", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d", "%m/%d/%y"):
                try:
                    return datetime.strptime(date_value, fmt).date()
                except ValueError:
                    continue
            raise ValueError(f"Unrecognized date format: {date_value}")
        elif isinstance(date_value, datetime):
            return date_value.date()
        elif isinstance(date_value, date):
            return date_value
        else:
            raise TypeError("Unsupported date type")


    def gregorian_to_jp(date_value):
        """Convert Gregorian date to Japanese era + formatted outputs."""
        dt = parse_date(date_value)
        # Determine Japanese era
        if dt >= date(2019, 5, 1):
            era = 'Reiwa'
            year = dt.year - 2018
        elif dt >= date(1989, 1, 8):
            era = 'Heisei'
            year = dt.year - 1988
        elif dt >= date(1926, 12, 25):
            era = 'Showa'
            year = dt.year - 1925
        else:
            era = 'Unknown'
            year = dt.year

        # Format outputs
        jp_year = str(year)
        month = f"{dt.month:02d}"
        day = f"{dt.day:02d}"
        joined_text = (f"{year:02d}.{month}.{day}")

        return era, jp_year, month, day, joined_text


    # Fields that need YYYYMMDD format
    assignment_date_fields = [
        "派遣期間 | Assignment period (自 | From) (yyyy/mm/dd)",
        "派遣期間 | Assignment period (至 | To) (yyyy/mm/dd)"
    ]
    def format_date_digits(date_value):
        """Convert date-like input into YYYY.MM.DD string."""
        dt = parse_date(date_value)
        return dt.strftime("%Y.%m.%d")
    
    # Fields that need postal code format
    post_code_fields = [
        "国内における住所の郵便番号 | Postal code",
        "派遣元事業所の郵便番号 | Zip code"
    ]

    def format_post_code(post_code_value):
        postcode=str(post_code_value).strip().replace("-","")
        if len(postcode)==7 and postcode.isdigit():
            return f"{postcode[:3]}-{postcode[3:]}"
        else:
        # Return original if it doesn't match expected format
            return post_code_value

    # --- Process PDF ---
    with fitz.open(template_pdf) as doc:
        for page in doc:
            for field in page.widgets():
                field_name = field.field_name
                if not field_name:
                    continue

                
                
        # Add this inside your existing loop, after determining `value` for normal text fields
                kana_column = '氏名 | ﾌﾘｶﾞﾅ | Name (in Kana)'
                japanese_fields = [
                    "国内における住所 | Address in Japan (in Japanese)",
                    "派遣元事業所の名称 | Home Company Name",
                    "所在地 | Office Address (in Japanese)",
                    "事業主氏名 | Name of representative",
                    kana_column,
                    "漢字氏名 | Name (in Kanji)",
                    "国内における住所 | Address in Japan (in Kana)",
                ]
                
                if field_name in japanese_fields and value:  # <<< UPDATED
                    text = str(data.get(field_name, '')).strip()
                    rect = field.rect  # <<< UPDATED: get field rectangle

                    # Create a TextWriter for the page
                    from fitz import TextWriter
                    n_chars = len(text)

                    # --- Progressive font size ---
                    min_chars, max_chars = 5, 30
                    max_font, min_font = 15, 7

                    if n_chars <= min_chars:
                        font_size = max_font
                    elif n_chars >= max_chars:
                        font_size = min_font
                    else:
                        # Linear interpolation
                        font_size = max_font - ((n_chars - min_chars) / (max_chars - min_chars)) * (max_font - min_font)


                    writer = TextWriter(page.rect)

                    x = rect.x0
                    y = rect.y0 + rect.height * 0.25  # adjust downward
                    writer.color = (0.09, 0.09, 0.09)  

                    # Append the text to the page at the widget's rectangle position
                    writer.append(
                        pos=(x,y),
                        text=str(data.get(field_name, '')).strip(),
                        fontsize=font_size,              # font size
                    )

                    # Write the text to the page
                    writer.write_text(page)  # <<< UPDATED: draws text directly onto the page

                    # Clear the form field so it disappears
                    field.field_value = ""   # <<< UPDATED
                    field.update()           # <<< UPDATED
                    continue                 # <<< UPDATED: skip normal field assignment

            # handle default selection
                
                if field_name in ['Default Selection']:
                    field.field_value = 'Yes'
                    field.update()
                    continue



                if field_name not in data and field_name not in [
                    'Era_Showa', 'Era_Heisei', 'Era_Reiwa', 'DOB_JP',
                    'Sex_Male', 'Sex_Female', 'Today_JPYear', 'Today_Month', 'Today_Day'
                ]:
                    continue



                # --- DOB handling ---
                dob_col_name = '生年月日 | Date of birth (yyyy/mm/dd)'
                if field_name in ['Era_Showa', 'Era_Heisei', 'Era_Reiwa', 'DOB_JP'] and dob_col_name in data:
                    era, jp_year, month, day, dob_text = gregorian_to_jp(data[dob_col_name])
                    if field_name == 'Era_Showa':
                        field.field_value = 'Yes' if era == 'Showa' else ''
                    elif field_name == 'Era_Heisei':
                        field.field_value = 'Yes' if era == 'Heisei' else ''
                    elif field_name == 'Era_Reiwa':
                        field.field_value = 'Yes' if era == 'Reiwa' else ''
                    elif field_name == 'DOB_JP':
                        field.field_value = dob_text
                    field.update()
                    continue

                # --- Sex handling ---
                if field_name in ['Sex_Male', 'Sex_Female'] and '性別 | Sex' in data:
                    sex_value = str(data['性別 | Sex']).strip()
                    if field_name == 'Sex_Male':
                        field.field_value = 'Yes' if sex_value == '男' else ''
                    elif field_name == 'Sex_Female':
                        field.field_value = 'Yes' if sex_value == '女' else ''
                    field.update()
                    continue

                # --- Today's Date handling ---
                if field_name in ["Today_JPYear", "Today_Month", "Today_Day"]:
                    era, jp_year, month, day, _ = gregorian_to_jp(date.today())
                    if field_name == "Today_JPYear":
                        field.field_value = jp_year
                    elif field_name == "Today_Month":
                        field.field_value = month
                    elif field_name == "Today_Day":
                        field.field_value = day
                    field.update()
                    continue
                    

                # --- Normal text / checkbox handling ---
                field_type = field.field_type
                value = str(data.get(field_name, '')).strip()


                if field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
                    value = str(data.get(field_name, ''))
                    field.field_value = 'Yes' if value in ['X', 'Yes'] else ''
                
                else:
                    if value in ['nan', 'None', '']:
                        field.field_value = ''
                    elif field_name in assignment_date_fields:
                        field.field_value = format_date_digits(value)
                    elif field_name in post_code_fields:
                        field.field_value = format_post_code(value)
                    else:
                        field.field_value = value    
                field.update()

        doc.save(output_pdf, deflate=True,incremental=False, clean=True)



# ===============================
# Streamlit UI
# ===============================
st.set_page_config(page_title="Multi-Country PDF Filler", layout="centered")
st.title("📄 Multi-Country CoC PDF Generator")
st.write("Upload an Excel file and a ZIP of country PDF templates. The app will generate filled PDFs per person and country.")
st.write("**Please ensure you use the exact template and follow the data format as noted in the Excel template.**")
# Upload Excel
excel_file = st.file_uploader("Upload Excel File", type=["xlsx"])

# Upload ZIP of PDF templates
zip_file = st.file_uploader("Upload ZIP of PDF Templates", type=["zip"])

# Generate PDFs
if excel_file and zip_file:
    if st.button("Generate PDFs"):
        try:
            # Save Excel to temp file
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_xlsx:
                tmp_xlsx.write(excel_file.read())
                excel_path = tmp_xlsx.name

            # Save ZIP to temp file and extract
            with tempfile.TemporaryDirectory() as tmp_zip_dir:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as tmp_zip_file:
                    tmp_zip_file.write(zip_file.read())
                    tmp_zip_path = tmp_zip_file.name

                with zipfile.ZipFile(tmp_zip_path, 'r') as zip_ref:
                    zip_ref.extractall(tmp_zip_dir)

                # Read Excel
                df = pd.read_excel(excel_path, dtype=str, header=2)

                output_files = []

                for index, row in df.iterrows():
                    data = row.to_dict()
                    country = str(data.get("Country", "US")).strip().upper()  # normalize

                    # Search for matching PDF template in extracted folder
                    def normalize_name(s):
                        return s.strip().lower().replace("_", " ")
                    
                    matches = []
                    excel_country = normalize_name(data.get("Country", "US"))
                    for f in glob.glob(os.path.join(tmp_zip_dir, "**/*.pdf"), recursive=True):
                        file_country = normalize_name(os.path.splitext(os.path.basename(f))[0])
                        if file_country == excel_country:
                            matches.append(f)
                    if not matches:
                        st.warning(f"Template for {country} not found for {data.get('漢字氏名 | Name (in Kanji)', 'Unknown')}")
                        continue
                    template_path = matches[0]

                    # Prepare output PDF path
                    full_name = data.get('漢字氏名 | Name (in Kanji)', f'User{index}')
                    year = 2025
                    output_pdf = os.path.join(tmp_zip_dir, f"{full_name} - {year} {country} CoC form.pdf")

                    # Fill the PDF
                    fill_and_flatten_pdf(template_path, output_pdf, data)
                    output_files.append(output_pdf)

                if not output_files:
                    st.error("No PDFs were generated. Please check your Excel and ZIP templates.")
                else:
                    # Create ZIP of all generated PDFs
                    final_zip_path = os.path.join(tmp_zip_dir, "Generated_PDFs.zip")
                    with zipfile.ZipFile(final_zip_path, "w") as zipf:
                        for f in output_files:
                            zipf.write(f, os.path.basename(f))

                    # Provide download button
                    with open(final_zip_path, "rb") as f:
                        st.success("✅ PDFs generated successfully!")
                        st.download_button(
                            label="Download All PDFs (ZIP)",
                            data=f,
                            file_name="Generated_PDFs.zip",
                            mime="application/zip"
                        )

        except Exception as e:
            st.error(f"An error occurred: {e}")
