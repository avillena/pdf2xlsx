import os
import argparse
from openpyxl import load_workbook
from img2table.document import PDF 
# Tesseract OCR
from img2table.ocr import TesseractOCR

tesseract_ocr = TesseractOCR(n_threads=1, lang="eng")

def pdf_to_xlsx(input_pdf, output_xls, ocr_tool, implicit_rows=False, borderless_tables=False, min_confidence=50):
    """
    Convert a PDF file to an Excel file using a specific OCR tool.

    :param input_pdf: Path to the input PDF file.
    :param output_xls: Path to the output Excel file.
    :param ocr_tool: OCR tool to use.
    :param implicit_rows: Boolean indicating whether to use implicit rows.
    :param borderless_tables: Boolean indicating whether to use borderless tables.
    :param min_confidence: Minimum confidence for OCR tool.
    """
    pdf = PDF(src=input_pdf)

    # Export to file
    pdf.to_xlsx(output_xls,
                ocr=ocr_tool,
                implicit_rows=implicit_rows,
                borderless_tables=borderless_tables,
                min_confidence=min_confidence)

def get_worksheet_info(output_xls):
    """
    Get and print information about worksheets in an Excel file.

    :param output_xls: Path to the Excel file.
    """
    for ws in load_workbook(output_xls):
        print(f"Worksheet {ws.title} : {len(tuple(ws.rows))} rows, {len(tuple(ws.rows)[0])} columns")

def main():
    # Definir tus variables aquí
    parser = argparse.ArgumentParser(description='Process a PDF file to an Excel file.')
    parser.add_argument('input_pdf', type=str, help='Path to the input PDF file.')
    args = parser.parse_args()

    input_pdf_path = args.input_pdf
    base_name = os.path.splitext(os.path.basename(input_pdf_path))[0]
    output_xls_path = f"{base_name}.xlsx"
    ocr_tool = tesseract_ocr # Asegúrate de que este objeto esté definido e importado correctamente

    pdf_to_xlsx(input_pdf_path, output_xls_path, ocr_tool)
    get_worksheet_info(output_xls_path)

if __name__ == "__main__":
    main()
