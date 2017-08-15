import gspread
from openpyxl import Workbook, load_workbook
from os import remove

import config.user_info as usr
from config.gsheets import client

# 2 sheets of data required
# contacts = client.open_by_url(usr.contacts_url).sheet1
# reimbs = client.open_by_url(usr.rmbs_url).sheet1

class Template:
    """
    Represents an Excel form template, that can additionally
    be filled with data.
    """

    def _save_file(self, wb, filename):
        """
        Saves file if filename is specified.
        """
        if filename:
            wb.save(filename)

    def get_empty(self, filename=False):
        """
        Returns empty template.
        Optionally saves as an Excel workbook if filename is
        specified.
        """
        wb = load_workbook(filename='template.xlsx')
        self._save_file(wb, filename)
        return wb

    def get_new(self, filename=False):
        """
        Returns new template with org_name, bookkeeper and
        treasurer variables filled.
        Optionally saves as an Excel workbook if filename is
        specified.
        """
        wb = load_workbook(filename = 'template.xlsx')
        ws = wb.active

        ws[usr.tmp_vars['org_name']] = usr.org_name
        ws[usr.tmp_vars['bookkeeper']] = usr.bookkeeper
        ws[usr.tmp_vars['treasurer']] = usr.treasurer

        self._save_file(wb, filename)
        return wb

    def get_completed(self, data_dict, filename=False):
        """
        Return completed template file given a dictionary of values
        to plug into the template.
        Optionally saves as an Excel workbook if filename is
        specified.

        data_dict: Dictionary of user specific data for each form, where
        keys are strings corresponding to usr.tmp_vars key names. 
        """
        wb = self.get_new()
        ws = wb.active

        for varname in data_dict:
            # only available data will be filled in
            ws[usr.tmp_vars[varname]] = data_dict[varname]

        self._save_file(wb, filename)
        return wb
