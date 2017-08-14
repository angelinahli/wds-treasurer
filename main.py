import gspread

from config/gsheets import client, rmbs_sht, contacts_sht 

# 2 sheets of data required
contacts = client.open(contacts_sht).sheet1
reimbs = client.open(rmbs_sht).sheet1