import logging
import openpyxl
from openpyxl.styles import colors
from openpyxl import Workbook
from openpyxl.styles import Font, Color, Alignment


def format_workbook(adjust_content=True, font=None, header=1,\
                    odd_bg=None, even_bg=None, thousand_separator=True):
    # print("Formatting workbook")
    logging.info('Format workbook')
    logging.info('adjust_content %r', adjust_content)
    logging.info('font %s', font)
    logging.info('header %d', header)
    logging.info('odd_bg %s', odd_bg)
    logging.info('even_bg %s', even_bg)
    logging.info('thousand_separator %r', thousand_separator)
    wb = openpyxl.load_workbook(filename='./res/input.xlsx')


    # Print current sheetnames
    print (wb.sheetnames)

    # Worksheet
    ws = wb.active

    # 1. Adjust content
    for col in ws.columns:
         max_length = 0
         column = col[0].column # Get the column name
         for cell in col:
             try: # Necessary to avoid error on empty cells
                 if len(str(cell.value)) > max_length:
                     max_length = len(cell.value)
             except:
                 pass
         #adjusted_width = (max_length + 2) * 1.2
         adjusted_width = (max_length + 2)
         ws.column_dimensions[column].width = adjusted_width
    # for row in ws.iter_rows():
    #     for cell in row:
    #         cell.style = Alignment(wrapText=True)


    # Change background<
    for row in ws.iter_rows():
        for cell in row:
            cell.font = Font(name='Ubuntu')
            #cell.font = Font(name='Ubuntu', color=colors.RED)

    # Save output to new filename
    wb.save(filename='./res/input_generated.xlsx')

    # Close workbook
    wb.close()


if __name__ == "__main__":
    FORMAT = "%(asctime)-7s | %(levelname)s  | %(message)s"
    logging.basicConfig(format=FORMAT, level=logging.DEBUG, datefmt='%H:%M:%S')
    # logging.debug('Debugging message')
    logging.info('My first step')
    # logging.warning('And this, too')
    format_workbook()
