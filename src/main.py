import logging
import openpyxl
from openpyxl.styles import colors
from openpyxl import Workbook
from openpyxl.styles import Font, Color, Alignment, PatternFill


class VarianzaFormat:

    class Options:
        adjust_content = True
        font = None
        header = 1
        header_bg = None
        odd_bg = None
        even_bg = None
        thousand_separator = True
        input_filepath = './res/input.xlsx'
        output_filepath = './res/input_generated.xlsx'

    def run(self, options):
        # print("Formatting workbook")
        logging.info('Format workbook')
        logging.info('adjust_content %r', options.adjust_content)
        logging.info('font :      %s', options.font)
        logging.info('header :    %d', options.header)
        logging.info('header_bg : %s', options.header_bg)
        logging.info('odd_bg :    %s', options.odd_bg)
        logging.info('even_bg :   %s', options.even_bg)
        logging.info('thousand_separator : %r', options.thousand_separator)
        wb = openpyxl.load_workbook(filename=options.input_filepath)

        # Select active Worksheet
        ws = wb.active

        # 1. Change cell font
        if options.font is not None:
            for row in ws.iter_rows():
                for cell in row:
                    cell.font = Font(name=options.font)

        # 1. Adjust content
        # See https://stackoverflow.com/questions/13197574/openpyxl-adjust-column-width-size/14450572
        if options.adjust_content:
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
            if options.header == row_number:
                bg = options.header_bg
            else:
                if row_number % 2 == 0:
                    bg = options.even_bg  # Even
                else:
                    bg = options.odd_bg  # Odd
            for cell in row:
                if bg is not None:
                    cell.fill = PatternFill(fgColor=bg, fill_type="solid")

        # 4. Add thousands separator
        if options.thousand_separator:
            for row in ws.iter_rows():
                for cell in row:
                    cell.number_format = u'#,##0.00'

        # Save output to new filename
        wb.save(filename=options.output_filepath)

        # Close workbook
        wb.close()


if __name__ == "__main__":
    FORMAT = "%(asctime)-7s | %(levelname)s  | %(message)s"
    logging.basicConfig(format=FORMAT, level=logging.DEBUG, datefmt='%H:%M:%S')
    logging.info('-- Starting --')
    # Set options for VarianzaFormat
    options = VarianzaFormat.Options()
    options.font = 'Consolas'
    options.header_bg = '0000FF'
    options.odd_bg = 'FF0000'
    options.even_bg = '00F700'

    # Run the class
    format = VarianzaFormat()
    format.run(options)
