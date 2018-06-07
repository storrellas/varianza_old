import logging
import openpyxl
from openpyxl.styles import colors
from openpyxl import Workbook
from openpyxl.styles import Font, Color, Alignment, PatternFill


def format_workbook(adjust_content=True, font=None, header=1, header_bg=None,
                    odd_bg=None, even_bg=None, thousand_separator=True):
    # print("Formatting workbook")
    logging.info('Format workbook')
    logging.info('adjust_content %r', adjust_content)
    logging.info('font :      %s', font)
    logging.info('header :    %d', header)
    logging.info('header_bg : %s', header_bg)
    logging.info('odd_bg :    %s', odd_bg)
    logging.info('even_bg :   %s', even_bg)
    logging.info('thousand_separator : %r', thousand_separator)
    wb = openpyxl.load_workbook(filename='./res/input.xlsx')

    # Select active Worksheet
    ws = wb.active

    # 1. Change cell font
    if font is not None:
        for row in ws.iter_rows():
            for cell in row:
                cell.font = Font(name=font)

    # 1. Adjust content
    # See https://stackoverflow.com/questions/13197574/openpyxl-adjust-column-width-size/14450572
    for col in ws.columns:
        max_length = 0
        column = col[0].column  # Get the column name

        for cell in col:
            try:  # Necessary to avoid error on empty cells
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except Exception:
                pass
        adjusted_width = (max_length + 2) * 2
        ws.column_dimensions[column].width = adjusted_width

    # 3. Update cell background (header, odd,even)
    for row in ws.iter_rows():

        bg = 'FFFFFF'
        row_number = row[0].row
        if header == row_number:
            bg = header_bg
        else:
            if row_number % 2 == 0:
                bg = even_bg  # Even
            else:
                bg = odd_bg  # Odd
        for cell in row:
            if bg is not None:
                cell.fill = PatternFill(fgColor=bg, fill_type="solid")

    # 4. Add thousands separator
    if thousand_separator:
        for row in ws.iter_rows():
            for cell in row:
                cell.number_format = u'#,##0.00'

    # Save output to new filename
    wb.save(filename='./res/input_generated.xlsx')

    # Close workbook
    wb.close()


if __name__ == "__main__":
    FORMAT = "%(asctime)-7s | %(levelname)s  | %(message)s"
    logging.basicConfig(format=FORMAT, level=logging.DEBUG, datefmt='%H:%M:%S')
    logging.info('-- Starting --')
    format_workbook(adjust_content=True, font='Consolas', header=1,
                    header_bg='0000FF', odd_bg='FF0000',
                    even_bg='00F700', thousand_separator=True)
