import gspread
from re import split

import config.user_info as usr
from config.gsheets import client

class Worksheet:

    def __init__(self, url, column_values, start_row, is_ordered=True):
        """
        url (string): the URL of a google sheet
        column_values (dict): a dictionary of the worksheets column names and
        their google sheets column indices
        start_row (int): the google sheets row index where the first line of
        data starts
        is_ordered (boolean): indicates whether data in this worksheet is filled
        from the starting row down
        """
        self.url = url
        self.column_values = column_values
        self.start_row = start_row
        self.is_ordered = is_ordered
        self.ws = client.open_by_url(self.url).sheet1


class UserWorksheet(Worksheet):

    def get_user_data(self):
        """
        returns dictionary of dictionaries for each user, with username as key
        and user data dictionary as values.
        """
        data_column_values = {
            var: value - 1 for var, value in self.column_values.items() \
            if var != "username"}
        user_dict = {}
        for row_num in range(self.start_row, self.ws.row_count + 1):
            row_data = self.ws.row_values(row_num)
            if self.is_ordered and not any(row_data):
                # stop checking rest of rows
                break
            username = row_data[self.column_values["username"] - 1]
            user_dict[username] = {
                var: row_data[data_column_values[var]] for var in 
                data_column_values}
        return user_dict

class ReimbursementWorksheet(Worksheet):

    def get_new_data_range(self):
        """
        returns a list of the row numbers that haven't yet been processed.
        """
        row_range = []
        for row_num in range(self.start_row, self.ws.row_count + 1):
            has_date = self.ws.cell(row_num, self.column_values["date"]).value
            if self.is_ordered and not has_date:
                break
            if not self.ws.cell(row_num, self.column_values["has_form"]).value:
                row_range.append(row_num)
        return row_range

    def get_rmbs_row_data(self, row_data):
        """
        row_data (list): list of data from a single row.
        returns all reimbursement data associated with specified row.
        """
        return {var: row_data[self.column_values[var] - 1] for var in
            self.column_values}

    def update_processed_rows(self, row_range):
        for row_num in row_range:
            self.ws.update_cell(row_num, self.column_values["has_form"], True)
        print "Row range: {} has been processed!".format(row_range)

