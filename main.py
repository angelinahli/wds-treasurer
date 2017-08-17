from os import makedirs
from os.path import basename
from shutil import rmtree
from smtplib import SMTP
from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.utils import COMMASPACE, formatdate

import config.user_info as usr
from new_forms import make_forms

OUTDIR = "output/user"

def clear_cache(outdir):
    """
    Removes all files in the output directory before generating
    new forms.
    """
    rmtree(outdir)
    makedirs(outdir)
    print "Cache cleared"

def send_attachments(files):
    """
    Given a list of form filepaths, will email those forms to a
    list of target addresses. 
    """
    msg = MIMEMultipart()
    msg["Subject"] = "WDS - New reimbursements"
    msg["From"] = usr.EMAIL_LOGIN + "@gmail.com"
    msg["To"] = COMMASPACE.join(usr.TARGET_ADDRESSES)
    msg["Date"] = formatdate(localtime=True)

    filenames = {file: basename(file) for file in files}
    body_text = ["The following files are attached:"]
    for file in filenames.values():
        body_text.append("\n*".format(file))
    body = MIMEText(''.join(body_text), 'plain')
    msg.attach(body)

    for file in filenames:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(open(file, "rb").read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition", 
            "attachment", 
            filename=filenames[file]
            )
        msg.attach(part)

    server = SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(usr.EMAIL_LOGIN, usr.EMAIL_PASSWORD)
    server.sendmail(msg["From"], msg["To"].split(COMMASPACE),msg.as_string())
    server.quit()
    print "Email sent"

def run(outdir):
    """
    Given the path of the output directory, will:
    (a) clear cache within that outdir;
    (b) generate forms for any new reimbursement entries and
        save them in the outdir;
    (c) email user the results
    """
    print "**Running program to generate reimbursement forms**"
    clear_cache(outdir)
    # make_forms returns form pathdirs
    forms = make_forms(outdir)
    send_attachments(forms)

run(OUTDIR)
