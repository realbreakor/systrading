import os
import logging
from dotenv import load_dotenv

load_dotenv()

APP_KEY = os.getenv("APP_KEY")
SECRET_KEY = os.getenv("SECRET_KEY")
ID = os.getenv('ID')
ID_PSWD = os.getenv('ID_PSWD')
PSWD = os.getenv('PSWD')
CERT = os.getenv('CERT')
MOUI_ID_PSWD = os.getenv('MOUI_ID_PSWD')
MOUI_PSWD = os.getenv('MOUI_PSWD')

BASE_DIR = os.path.expanduser("~/Documents/PycharmProjects/systrading/")
FUN_DIR = os.path.expanduser("~/")

logging.basicConfig(
    filename=BASE_DIR + 'order_log.txt',
    level=logging.INFO,
    format='%(asctime)s - %(message)s'
)
