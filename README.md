import fitz  # PyMuPDF
import pandas as pd

# Define a function to extract text from a specific page of the PDF
def extract_text_from_page(pdf_path, page_number):
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_number)
    text = page.get_text()
    return text

# Define a function to process the text and structure the data
def process_text(text):
    lines = text.split('\n')
    data = []
    for line in lines:
        if line.strip():
            # This is a simplified example. Needs adjustment based on the exact format
            parts = line.strip().split(maxsplit=6)
            if len(parts) == 7:
                data.append(parts)
    return data

# Define a function to create a pandas DataFrame
def create_dataframe(data):
    columns = ["Descripción", "Cantidad", "Unidad", "Rendimiento", "Precio Unitario", "Costo Total", "Mano de Obra"]
    df = pd.DataFrame(data, columns=columns)
    return df

# Define a function to verify and extract text from a range of pages
def verify_and_extract(pdf_path, sections):
    doc = fitz.open(pdf_path)
    total_pages = doc.page_count
    extracted_data = {}

    for section, page in sections.items():
        if page < total_pages:
            text = extract_text_from_page(pdf_path, page)
            data = process_text(text)
            if data:
                extracted_data[section] = data
        else:
            print(f"Page {page} for section {section} is out of range.")
    return extracted_data

# Main function to process the PDF and create the Excel file
def main():
    pdf_path = "/path/to/your/pdf/computos-y-presupuestos-mario-chandias-comprimido.pdf"
    output_excel = "/path/to/save/computos_y_presupuestos.xlsx"

    # Adjusted pages based on PDF's table of contents
    sections = {
        "Movimiento de Tierra": 27,
        "Albañilería": 51,
        "Hormigón Armado": 125,
        "Carpintería Metálica y de Madera": 173,
        "Pintura": 195,
        "Revestimientos y Sillería": 245,
        "Instalaciones Complementarias": 299
    }

    # Extract and process data
    extracted_data = verify_and_extract(pdf_path, sections)

    # Create an Excel writer
    writer = pd.ExcelWriter(output_excel, engine='xlsxwriter')

    for section, data in extracted_data.items():
        df = create_dataframe(data)
        df.to_excel(writer, sheet_name=section, index=False)

    writer.save()

# Execute the main function
main()- 👋 Hi, I’m @MarcosAO2401
- 👀 I’m interested in ...
- 🌱 I’m currently learning ...
- 💞️ I’m looking to collaborate on ...
- 📫 How to reach me ...
- 😄 Pronouns: ...
- ⚡ Fun fact: ...

<!---
MarcosAO2401/MarcosAO2401 is a ✨ special ✨ repository because its `README.md` (this file) appears on your GitHub profile.
You can click the Preview link to take a look at your changes.
--->
