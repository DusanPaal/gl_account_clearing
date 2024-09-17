"""
The 'biaMail.py' module creates and sends emails directly via a SMTP server.

Version history:
1.0.20220125 - Initial version.
"""

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from logging import Logger, getLogger
from os.path import isfile
import smtplib
from smtplib import SMTPConnectError
from typing import Union

_logger: Logger = getLogger("master")

def send_message(sender: str, subj: str, body: str, to: Union[list, str], cc: Union[list, str] = None, att_path: str = None) -> bool:
    """
    Creates and sends an email.

    Params:
        sender: Email address of the sender.
        subj: Subject of the email message.
        body: Body of the message in HTML format.
        to: Email addresses of message recipients.
        cc: Email Addresses of message carbon copy recipients.
        att_path: Path ot the attachment file.

    Returns: True if message was sent, False if not.
    """

    assert type(sender) is str, "Argument 'sender' has invalid type!"
    assert "@" in sender, "Invalid 'sender' argument!"
    assert type(subj) is str, "Argument 'subj' has invalid type!"
    assert type(body) is str, "Argument 'body' has invalid type!"
    assert type(to) in (list, str), "Argument 'to' has invalid type!"
    assert len(to) != 0, "Invalid 'to' argument!"
    assert cc is None or type(cc) in (list, str), "Argument 'cc' has invalid type!"

    if att_path is not None:
        assert type(att_path) is str, "Argument 'att_path' has invalid type!"
        assert isfile(att_path), f"Attachment '{att_path}' not found!"

    EMPTY_STR = ""

    # email debugging constants
    DEBUG_OFF = 0
    DEBUG_VERBOSE = 1
    DEBUG_VERBOSE_TIMESTAMPED = 2

    att_file_paths = []
    to_mails = []
    cc_mails = []

    if type(to) is list:
        to_mails = to
    elif type(to) is str:
        to_mails.append(to)

    if type(cc) is list:
        cc_mails = cc
    elif type(cc) is str:
        cc_mails.append(cc)

    recips = to_mails + cc_mails

    if att_path is not None:
        att_file_paths.append(att_path)

    email = MIMEMultipart()
    email["Subject"] = subj
    email["From"] = sender
    email["To"] = ";".join(to_mails)
    email["Cc"] = ";".join(cc_mails)
    email.attach(MIMEText(body, "html"))

    if att_path is not None:

        _logger.debug(f" Attaching file: {att_path}")

        att = open(att_path, "rb")
        payload = att.read()
        att.close()

        # The content type "application/octet-stream" means
        # that a MIME attachment is a binary file
        part = MIMEBase("application", "octet-stream")
        part.set_payload(payload)
        encoders.encode_base64(part) # Encode to base64

        # get file name
        file_name = att_path.split("\\").pop()

        # Add header
        part.add_header(
            "Content-Disposition",
            f"attachment; filename = {file_name}"
        )

        # Add attachment to your message and convert it to string
        email.attach(part)

    _logger.debug(f" Serializing message ...")
    text = email.as_string()

    host_name = "intrelay.ledvance.com"
    port_num = 25
    conn_timeout = 15

    _logger.debug(" Setting up SMTP connection using params: "
        f"Host = {host_name}; Port = {port_num}; Timeout = {conn_timeout}"
    )

    try:
        smtp_conn = smtplib.SMTP(host_name, port_num, timeout = conn_timeout)
    except SMTPConnectError:
        _logger.error(" Failed to setup SMTP connection!")
        return False
    except TimeoutError:
        _logger.error(" Waiting for SMTP connection timed out!")
        return False

    smtp_conn.set_debuglevel(DEBUG_OFF)

    err_msg = None

    _logger.debug(" Sending message ...")

    try:
        send_errs = smtp_conn.sendmail(sender, recips, text)
    except Exception as exc:
        err_msg = exc
    finally:
        smtp_conn.quit()

    if err_msg is None and len(send_errs) != 0:
        failed_recips = ';'.join(send_errs.keys())
        _logger.warning(f"Message sent, but the following recipients failed to receive the message: {failed_recips}")
    elif err_msg is not None:
        _logger.error(f"Sending failed. Reason: {err_msg}")
        return False

    return True
