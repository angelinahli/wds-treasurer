from os.path import basename
from smtplib import SMTP
from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.utils import COMMASPACE, formatdate

import config.user_info as usr
from temp.temp_email import temp_body

def send_forms(file_data, testing=False):
    """
    Given the output of new_forms.make_forms, will email those forms to a
    list of target addresses.
    Each line of file_data is structured as [pathname, {check_value: bool}]
    """
    if file_data:
        from_email = usr.EMAIL_LOGIN + "@gmail.com"
        to_email = usr.TARGET_ADDRESSES
        if testing:
            to_email = usr.ADMIN_ADDRESS

        msg = MIMEMultipart()
        msg["Subject"] = "WDS - New reimbursements"
        msg["From"] = from_email
        msg["To"] = COMMASPACE.join(to_email)
        msg["Date"] = formatdate(localtime=True)

        body_text = [temp_body]
        for file in file_data:
            filename = basename(file[0])
            text = ["- {}".format(filename)]
            warnings = file[1]
            if any(warnings.values()):
                msgs = " | warnings: {}".format(
                    ", ".join([val for val in warnings if warnings[val]])
                    )
                text.append(msgs)
            text.append("\n")
            body_text.append("".join(text))
            
        body = MIMEText(''.join(body_text), 'plain')
        msg.attach(body)

        paths = [file[0] for file in file_data]
        filenames = {pathname: basename(pathname) for pathname in paths}

        for path in filenames:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(open(path, "rb").read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition", 
                "attachment", 
                filename=filenames[path]
                )
            msg.attach(part)

        server = SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(usr.EMAIL_LOGIN, usr.EMAIL_PASSWORD)
        server.sendmail(from_email, to_email, msg.as_string())
        server.quit()
        print "Email sent"