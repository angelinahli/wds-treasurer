import gspread
from re import split

import config.user_info as usr
from config.gsheets import client
from program.template import FormTemplate

# Get user data

def get_user_dict():
    """
    returns dictionary of dictionaries for each user (student member),
    with username as key, and user's dictionary data as values.
    """
    cols = usr.contacts_cols
    contacts = client.open_by_url(usr.CONTACTS_URL).sheet1
    user_dict = {}
    for row_num in range(usr.CONTACTS_ROWSTART, contacts.row_count+1):
        data = contacts.row_values(row_num)
        if not any(data):
            # assumes rows are filled in from top down
            break
        user_dict[ data[cols['username'] - 1] ] = {
            var: data[cols[var] - 1] for var in [
                'stu_name', 'stu_id', 'stu_unit_box'
            ]
        }
    return user_dict

def get_new_row_range(ws):
    """
    given a gspread worksheet representing the form accepting new
    reimbursements, returns a list of the row range that have not
    yet been processed.
    """
    row_range = []
    for row_num in range(usr.RMBS_ROWSTART, ws.row_count+1):
        # if timestamp is missing, data assumed to be missing
        if not ws.cell(row_num, usr.rmbs_cols['date']).value:
            break
        # has_form is None if no form has been created yet
        if not ws.cell(row_num, usr.rmbs_cols['has_form']).value:
            row_range.append(row_num)
    return row_range

def update_rows(ws, row_range):
    """
    given a worksheet and a list of row range, updates rows to indicate a 
    form has been created for each of those rows.
    """
    for row_num in row_range:
        ws.update_cell(row_num, usr.rmbs_cols['has_form'], True)
    print "Rows {} have been processed!".format(row_range)

def get_rmbs_dict(row_data):
    """
    converts a list of reimbursements data for a given row to a dictionary
    of that data.
    """
    return {var: row_data[usr.rmbs_cols[var] - 1] for var in usr.rmbs_cols}

def get_student_data(student_usernames, user_data):
    """
    student_names = a string
    converts a delimited list of student usernames into a list of dictionaries
    of students and their contact information.
    """
    usernames = split('\s*,\s*', student_usernames)
    try:
        student_data = [user_data[user] for user in usernames]
    except KeyError:
        error_msg = "A user has not submitted contact info."
        raise ValueError(error_msg)
    return student_data

def format_data_dict(data_dict):
    """
    helper function for get_data_dict
    """
    data_dict['evt_cat'] = data_dict['evt_purpose']
    # assume funding comes from SOFC
    data_dict['evt_fund'] = 'SOFC'
    for value in ['evt_amt', 'evt_num']:
        current_value = data_dict[value]
        if current_value:
            data_dict[value] = int(current_value)
    data_dict['stu_unit_box'] = 'Unit {}'.format(data_dict['stu_unit_box'])
    return data_dict

def get_data_dict(ws, row_number, user_data):
    """
    given a gspread worksheet of the form accepting new reimbursements,
    the row number to access and a dictionary of data for each user,
    returns a dictionary of all data associated with that row of reimbursements.
    """
    rmbs_data = get_rmbs_dict(ws.row_values(row_number))
    username = rmbs_data['username']
    try:
        user_vals = dict(user_data[username], **rmbs_data)
    except KeyError:
        error_msg = "User {} has not submitted contact info.".format(username)
        raise ValueError(error_msg)
    data_dict = {var: user_vals[var] for var in usr.tmp_vars if var in user_vals}
    
    other_student_data = False
    other_students = user_vals['other_stu']
    if other_students:
        other_student_data = get_student_data(other_students, user_data)
    return format_data_dict(data_dict), other_student_data
    
def get_rmbs_data(ws, row_range):
    """
    given a gspread worksheet representing the form accepting new 
    reimbursements and list of row range, returns a list of dictionaries 
    that can be used to create spreadsheets.
    """
    user_data = get_user_dict()
    return [get_data_dict(ws, row, user_data) for row in row_range]

def make_other_students_file(filename, student_data):
    """
    If other students were associated with this reimbursement, will generate
    a document with the names of all students associated with this reimbursement.
    Returns the filepath associated with this file.
    """
    filepath = filename + ".txt"
    print "Generating file {}".format(filepath)
    with open(filepath, "w") as students_file:
        students_file.write("This reimbursement paid for the following students (Banner ID included in brackets for each):")
        for student in student_data:
            line = "\n- {name} ({id})".format(
                    name=student['stu_name'],
                    id=student['stu_id'])
            students_file.write(line)
    return filepath

# Make forms
def make_forms(outdir):
    """
    when run, will automatically generate forms for all non
    processed reimbursement entries, and save those forms in output/user.
    returns data for the forms generated for later use.
    """
    rmbs = client.open_by_url(usr.RMBS_URL).sheet1
    row_range = get_new_row_range(rmbs)
    print "Loading data for rows", row_range
    all_data = get_rmbs_data(rmbs, row_range)    
    filepaths = []
    for row_data in all_data:
        # this is the dictionary of data needed for a form
        data_dict = row_data[0]
        # date structured as: "%m/%d/%Y %H:%M:%S"
        date = data_dict['date'].split()[0].replace('/', '')
        # name structured as "Angelina Li"
        name = data_dict['stu_name'].split()[0].lower()
        filename = "{outdir}/{name}_{date}".format(
                    outdir=outdir,
                    name=name,
                    date=date)
        filepath = filename + ".xlsx"
        print "Generating form {}".format(filepath)
        FormTemplate(filepath=filepath).get_completed(data_dict)
        filepaths.append(filepath)
        
        # if other students were associated with this reimbursement, this is a
        # list of the contact data dictionaries associated with those students.
        other_students = row_data[1]
        if other_students:
            other_students_filepath = make_other_students_file(
                filename, 
                other_students)
            filepaths.append(other_students_filepath)

    update_rows(rmbs, row_range)
    return filepaths