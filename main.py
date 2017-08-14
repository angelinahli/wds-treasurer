import gspread

import config.user_info as usr
from config.gsheets import client

# 2 sheets of data required
contacts = client.open(usr.contacts_sht).sheet1
reimbs = client.open(usr.rmbs_sht).sheet1