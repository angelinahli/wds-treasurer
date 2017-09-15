import schedule
import time
from os import makedirs
from shutil import rmtree

from config.user_info import OUTDIR
from email_forms import Email
from new_forms import make_files


def clear_cache(outdir):
    """
    Removes all files in the output directory before generating
    new forms.
    """
    rmtree(outdir)
    makedirs(outdir)
    print "Cache cleared"

def run(outdir, testing_mode=False):
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
    filepaths = make_files(outdir)
    Email(filepaths, testing_mode).send_files("WDS - New Reimbursements")
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

run(OUTDIR, testing_mode=True)