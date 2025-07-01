import logging

RESET = "\033[0m"
DEBUG_COLOR = "\033[90m" # Grey
INFO_COLOR = "\033[36m"  # Cyan
WARNING_COLOR = "\033[33m" # Yellow
ERROR_COLOR = "\033[31m"  # Red
CRITICAL_COLOR = "\033[41m\033[37m" # White text on Red background

class ColoredFormatter(logging.Formatter):
    FORMATS = {
        logging.DEBUG: DEBUG_COLOR + '%(asctime)s - %(name)s - %(levelname)s - %(message)s' + RESET,
        logging.INFO: INFO_COLOR + '%(asctime)s - %(name)s - %(levelname)s - %(message)s' + RESET,
        logging.WARNING: WARNING_COLOR + '%(asctime)s - %(name)s - %(levelname)s - %(message)s' + RESET,
        logging.ERROR: ERROR_COLOR + '%(asctime)s - %(name)s - %(levelname)s - %(message)s' + RESET,
        logging.CRITICAL: CRITICAL_COLOR + '%(asctime)s - %(name)s - %(levelname)s - %(message)s' + RESET
    }

    def format(self, record):
        log_fmt = self.FORMATS.get(record.levelno)
        formatter = logging.Formatter(log_fmt)
        return formatter.format(record)
    
logger = logging.getLogger("ExcelExtract")

handler = logging.StreamHandler()
formatter = ColoredFormatter()
handler.setFormatter(formatter)

if logger.handlers:
    for handler in list(logger.handlers):
        logger.removeHandler(handler)

logger.addHandler(handler)

def addFileHandler(logfile):
    file_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    file_handler = logging.FileHandler(logfile, mode='w', encoding='utf-8')
    file_handler.setFormatter(file_formatter)
    logger.addHandler(file_handler)
