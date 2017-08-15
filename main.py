import gspread
from openpyxl import Workbook, load_workbook

import config.user_info as usr
from config.gsheets import client

# 2 sheets of data required
contacts = client.open_by_url(usr.contacts_url).sheet1
reimbs = client.open_by_url(usr.rmbs_url).sheet1

class Template:
    """
    Returns an Excel form, with the option of being filled.
    """

    def __init__(self):
        """
        template.xlsx is a clean Excel form file,
        tmp_vars is a dictionary of cell references that need to be filled in
        the template.
        """
        self.empty = load_workbook(filename = 'template.xlsx')
        self.tmp_vars = usr.tmp_vars

        self.org_name = usr.org_name
        self.bookkeeper = usr.bookkeeper
        self.treasurer = usr.treasurer

    def get_empty(self):
        """
        returns empty template.
        """
        return self.empty

    def get_new(self):
        """
        returns new template with org_name, bookkeeper and
        treasurer variables filled.
        """
        wb = self.empty # would like for this to be self.get_empty() ///// not sure how to do that yet
        ws = wb.active

        ws[self.tmp_vars['org_name']] = self.org_name
        ws[self.tmp_vars['bookkeeper']] = self.bookkeeper
        ws[self.tmp_vars['treasurer']] = self.treasurer

        return wb



