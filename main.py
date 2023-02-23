import pandas as pd
import glob
from fpdf import fpdf
from pathlib import Path
from datetime import date

_debug_ = True
if __name__ == '__main__':
    field_widths=(22, 55, 40, 35, 10)
    field_justify = ('L', 'L', 'C', 'C', 'L')
    # Get the list of files to convert
    filepaths = glob.glob('invoices/*.xlsx')

    if _debug_:
        print(filepaths)

    for filepath in filepaths:
        # Read the data
        df = pd.read_excel(filepath)
        total = 0
        field = 0

        if _debug_:
            print(df)

        # Convert to PDF
        pdf = fpdf.FPDF(orientation='P', unit='mm', format='A4')
        pdf.add_page()
        pdf.set_font(family='Arial', size=12)

        inv_no, ry_date = Path(filepath).name.removesuffix('.xlsx').split(sep='-')
        y, m, d = ry_date.split('.')
        ry_date = date(int(y), int(m), int(d)).strftime('%d %B, %Y')
        pdf.cell(w=0, h=12, align='L', ln=1, txt=f'Invoice Number: {inv_no}')
        pdf.cell(w=0, h=12, align='L', ln=1, txt=f'Date: {ry_date}')

        for header, w in zip(df.columns.values, field_widths):
            col_header = header.replace('_', ' ').title()
            if _debug_:
                print(col_header)

            pdf.cell(w=w, h=12, align='L', ln=0, txt=col_header)
        pdf.ln(h=5)

        for index, row in df.iterrows():
            if _debug_:
                print(index, row.values)

            for field, w, align in zip(row, field_widths, field_justify):
                pdf.cell(w=w, h=10, align=align, ln=0, txt=str(field))

            total = total + float(field)
            pdf.ln(h=5)

        pdf.cell(w=0, h=14, align='L', txt=f'Total = {total} euros')

    # Write the output
        outfile = str(Path('PDFs', Path(filepath).name).with_suffix('.pdf'))
        if _debug_:
            print(outfile)
        pdf.output(outfile)

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
