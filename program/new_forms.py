import gspread

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
    for row in range(usr.CONTACTS_ROWSTART, contacts.row_count+1):
        data = contacts.row_values(row)
        if not any(data):
            break
        user_dict[ data[cols['username'] - 1] ] = {
            'stu_name': data[cols['stu_name'] - 1],
            'stu_id': data[cols['stu_id'] - 1],
            'stu_unit_box': data[cols['stu_unit_box'] - 1]
        }
    return user_dict

def get_new_row_range(ws):
    """
    given a gspread worksheet representing the form accepting new
    reimbursements, returns a list of the row range that have not
    yet been processed.
    """
    row_range = []
    for row in range(usr.RMBS_ROWSTART, ws.row_count+1):
        # if timestamp is missing, data assumed to be missing
        if not ws.cell(row, usr.rmbs_cols['date']).value:
            break
        # has_form is None if no form has been created yet
        if not ws.cell(row, usr.rmbs_cols['has_form']).value:
            row_range.append(row)
    return row_range

def update_rows(ws, row_range):
    """
    given a worksheet and a list of row range, updates rows to indicate a 
    form has been created for each of those rows.
    """
    for row in row_range:
        ws.update_cell(row, usr.rmbs_cols['has_form'], True)
    print "Rows {} have been processed!".format(row_range)

def get_rmbs_dict(row_data):
    """
    converts a list of reimbursements data for a given row to a dictionary
    of that data.
    """
    return {var: row_data[usr.rmbs_cols[var] - 1] for var in usr.rmbs_cols}

def get_data_dict(ws, row_number, user_data):
    """
    given a gspread worksheet of the form accepting new reimbursements,
    the row number to access and a dictionary of data for each user,
    returns a dictionary of all data associated with that row of reimbursements.
    """
    rmbs_data = get_rmbs_dict(ws.row_values(row_number))
    username = rmbs_data['username']
    try:
        user_values = dict(user_data[username], **rmbs_data)
    except KeyError:
        raise ValueError("User {} has not submitted their info to the contacts form".format(username))
    data_dict = {var: user_values[var] for var in usr.tmp_vars if var in user_values}
    if ' for tournament' in data_dict['evt_purpose']:
        data_dict['evt_cat'] = data_dict['evt_purpose'].replace(' for tournament', '')
    if data_dict['evt_cat'] == 'Transportation':
        data_dict['evt_fund'] = 'SOFC'
    for value in ['evt_amt', 'evt_num']:
        current_value = data_dict[value]
        if current_value:
            data_dict[value] = int(current_value)
    data_dict['stu_unit_box'] = 'Unit {}'.format(data_dict['stu_unit_box'])
    return data_dict

def get_rmbs_data(ws, row_range):
    """
    given a gspread worksheet representing the form accepting new 
    reimbursements and list of row range, returns a list of dictionaries 
    that can be used to create spreadsheets.
    """
    user_data = get_user_dict()
    return [get_data_dict(ws, row, user_data) for row in row_range]

# Make forms

def make_forms(outdir):
    """
    when run, will automatically generate forms for all non
    processed reimbursement entries, and save those forms in output/user.
    returns the filepaths of the forms generated for later use.
    """
    rmbs = client.open_by_url(usr.RMBS_URL).sheet1
    row_range = get_new_row_range(rmbs)
    print "Loading data for rows", row_range
    all_data = get_rmbs_data(rmbs, row_range)
    filepaths = []
    for data_dict in all_data:
        # date structured as: "%m/%d/%Y %H:%M:%S"
        date = data_dict['date'].split()[0].replace('/', '')
        # name structured as "Angelina Li"
        name = data_dict['stu_name'].split()[0].lower()
        filepath = "{outdir}/{name}_{date}.xlsx".format(
                    outdir=outdir,
                    name=name,
                    date=date)
        print "Generating form {}".format(filepath)
        filepaths.append(filepath)
        FormTemplate(filepath=filepath).get_completed(data_dict)
    update_rows(rmbs, row_range)
    return filepaths