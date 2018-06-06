import logging
import openpyxl


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
