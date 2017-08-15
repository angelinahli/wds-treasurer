from openpyxl import load_workbook

import config.user_info as usr
from config.gsheets import client

class FormTemplate:
    """
    Represents an Excel form template, that can additionally
    be filled with data.
    """

    def _save_file(self, wb, filepath):
        """
        Saves file if filepath is specified.
        """
        if filepath:
            wb.save(filepath)

    def get_empty(self, filepath=None):
        """
        Returns empty template.
        Optionally saves as an Excel workbook if filepath is
        specified.
        """
        wb = load_workbook(filename='template.xlsx')
        self._save_file(wb, filepath)
        return wb

    def get_new(self, filepath=None):
        """
        Returns new template with org_name, bookkeeper,
        treasurer and address stub variables filled.
        Optionally saves as an Excel workbook if filepath is
        specified.
        """
        wb = load_workbook(filename='template.xlsx')
        ws = wb.active

        ws[usr.tmp_vars['org_name']] = usr.org_name
        ws[usr.tmp_vars['bookkeeper']] = usr.bookkeeper
        ws[usr.tmp_vars['treasurer']] = usr.treasurer
        ws[usr.tmp_vars['address']] = usr.address

        self._save_file(wb, filepath)
        return wb

    def get_completed(self, data_dict, filepath=None):
        """
        Return completed form given a dictionary of values
        to plug into the template.
        Optionally saves as an Excel workbook if filepath is
        specified.

        data_dict: Dictionary of user specific data for each form, where
        keys are strings corresponding to usr.tmp_vars key names. 
        """
        wb = self.get_new()
        self._save_file(wb, 'testing.xlsx')
        ws = wb.active
        for varname in data_dict:
            # only available data will be filled in
            ws[usr.tmp_vars[varname]] = data_dict[varname]
        self._save_file(wb, filepath)
        return wb