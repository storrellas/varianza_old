import logging


def format_workbook():
    #print("Formatting workbook")
    logging.info('Format workbook')

if __name__ == "__main__":
    FORMAT = "%(asctime)-7s | %(levelname)s  | %(message)s"
    logging.basicConfig(format=FORMAT, level=logging.DEBUG, datefmt='%H:%M:%S')
    #logging.debug('Debugging message')
    logging.info('My first step')
    #logging.warning('And this, too')
    format_workbook()
