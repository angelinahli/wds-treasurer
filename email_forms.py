from os.path import basename
from smtplib import SMTP
from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.utils import COMMASPACE, formatdate

import config.user_info as usr
from temp.temp_email import temp_body

class Email:
    """
    Represents an email that contains all the files created by this
    program.
    """

    def __init__(self, files, testing_mode=False):
        self.files = files
        self.testing_mode = testing_mode

    def _get_body(self, filenames):
        """
        Defines a new email body and returns it as a MIMEText object.
        Will only be called if there is at least one file in self.files.
        temp_body is a string generic email header.
        """
        body_text = []
        for file in filenames.values():
            body_text.append("- {}\n".format(file))
        return MIMEText(temp_body + ''.join(body_text), 'plain')

    def _get_multipart_email(self, subject, filenames):
        """
        Defines and returns a MIMEMultipart email object.
        subject: String subject.
        """
        self.from_email = usr.EMAIL_LOGIN + "@gmail.com"
        self.to_email = usr.ADMIN_ADDRESS if self.testing_mode \
                        else usr.TARGET_ADDRESSES

        msg = MIMEMultipart()
        msg["Subject"] = subject
        msg["From"] = self.from_email
        msg["To"] = COMMASPACE.join(self.to_email)
        msg["Date"] = formatdate(localtime=True)

        body = self._get_body(filenames)
        msg.attach(body)
        
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

        return msg

    def send_files(self, subject):
        """
        Defines a server and sends files if they exist. (if they don't exist,
        an empty list will be passed to the file parameter of this class)
        """
        file_attachments = self.files

        if file_attachments:
            filenames = {pathname: basename(pathname) for pathname in 
                         file_attachments}
            multipart_email = self._get_multipart_email(subject, filenames)

            server = SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(usr.EMAIL_LOGIN, usr.EMAIL_PASSWORD)
            server.sendmail(
                self.from_email, 
                self.to_email, 
                multipart_email.as_string())
            server.quit()

            print "Email sent"