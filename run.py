import schedule
import time
from os import makedirs
from shutil import rmtree

from config.user_info import OUTDIR
from program.email_forms import send_forms
from program.new_forms import make_forms

def clear_cache(outdir):
    """
    Removes all files in the output directory before generating
    new forms.
    """
    rmtree(outdir)
    makedirs(outdir)
    print "Cache cleared"

def run(outdir, testing=False):
    """
    Given the path of the output directory, will:
    (a) clear cache within that outdir;
    (b) generate forms for any new reimbursement entries and
        save them in the outdir;
    (c) email user the results
    If testing is set to true, only emails program admin.
    """
    start = time.time()
    print "**Running program to generate reimbursement forms**"
    clear_cache(outdir)
    # make_forms returns form pathdirs
    forms = make_forms(outdir)
    send_forms(forms, testing)
    print "Runtime:", time.time() - start

def scheduled_run(outdir, exec_time):
    """
    Given a string of the time to execute the function,
    will execute run() every week at approximately that time.
    """
    schedule.every().tuesday.at(exec_time).do(run(outdir))
    while True:
        schedule.run_pending()
        time.sleep(300)

run(OUTDIR, testing=True)