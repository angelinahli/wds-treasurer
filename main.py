import gspread

import config.user_info as usr
from config.gsheets import client

# 2 sheets of data required
contacts = client.open_by_url(usr.contacts_url).sheet1
reimbs = client.open_by_url(usr.rmbs_url).sheet1