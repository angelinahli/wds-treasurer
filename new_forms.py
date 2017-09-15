import gspread

import config.user_info as usr
from file_templates import Form, StudentsFile
from worksheets import ReimbursementWorksheet, UserWorksheet

def _get_all_row_data(rmbs_sheet, row_num, user_data):
    """
    rmbs_sheet (ReimbursementWorksheet): a Worksheet object that contains
    reimbursement information in each row.
    row_num (int): the current row whose data you want.
    user_data (dict): dictionary of dictionaries mapping usernames to user
    data dictionaries.
    returns all data associated with a specified row.
    """
    rmbs_data = rmbs_sheet.get_rmbs_row_data(rmbs_sheet.ws.row_values(row_num))
    username = rmbs_data["username"]
    try:
        row_data = dict(user_data[username], **rmbs_data)
    except KeyError:
        error_msg = "User {} has not submitted contact info.".format(
            username)
        raise KeyError(error_msg)
    return row_data

def get_all_data(rmbs_sheet, row_range, user_data):
    """
    rmbs_sheet (ReimbursementWorksheet): a Worksheet object that contains
    reimbursement information in each row.
    user_data (dict): dictionary of dictionaries mapping usernames to user
    data dictionaries.
    returns all data associated with each new row of reimbursements.
    """
    return [ _get_all_row_data(rmbs_sheet, row_num, user_data) for row_num in 
            row_range]

def make_files(outdir):
    """
    """
    rmbs_sheet = ReimbursementWorksheet(
        usr.RMBS_URL,
        usr.rmbs_cols,
        usr.RMBS_ROWSTART)
    row_range = rmbs_sheet.get_new_data_range()
    print "Loading data for rows", row_range

    user_data = UserWorksheet(
        usr.CONTACTS_URL,
        usr.contacts_cols,
        usr.CONTACTS_ROWSTART).get_user_data()
    all_data = get_all_data(rmbs_sheet, row_range, user_data)
    filepaths = []

    for row_data in all_data:
        # first generate a filename
        # date structured as: "%m/%d/%Y %H:%M:%S"
        date = row_data['date'].split()[0].replace('/', '')
        # name structured as "Angelina Li"
        name = row_data['stu_name'].split()[0].lower()
        filename = "{outdir}/{name}_{date}".format(
                    outdir=outdir,
                    name=name,
                    date=date)
        # then check to see if there are other students associated with
        # the reimbursement
        other_student_usernames = row_data["other_stu"]
        if other_student_usernames:
            student_file = StudentsFile(filepath=filename + ".txt")
            student_file.get_new_file(other_student_usernames, user_data)
            filepaths.append(student_file.filepath)
        # then make form
        form_data = {var: row_data[var] for var in usr.tmp_vars 
                     if var in row_data}
        rmbs_form = Form(filepath=filename + ".xlsx")
        rmbs_form.get_completed(form_data)
        filepaths.append(rmbs_form.filepath)

    rmbs_sheet.update_processed_rows(row_range)
    return filepaths