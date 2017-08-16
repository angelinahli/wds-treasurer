import smtplib
from email.mime.text import MIMEText
from os import makedirs
from shutil import rmtree

import config.user_info as usr
from new_forms import make_forms

def clear_cache(outdir):
    """
    Removes all files in the output directory before generating
    new forms.
    """
    rmtree(outdir)
    makedirs(outdir)
    print "Cache cleared"

def run(outdir):
    """
    Given the path of the output directory, will:
    (a) clear cache within that outdir;
    (b) generate forms for any new reimbursement entries and
        save them in the outdir;
    (c) email user the results
    """
    print "**Running program to generate reimbursement forms**\n"
    clear_cache(outdir)
    make_forms(outdir)

