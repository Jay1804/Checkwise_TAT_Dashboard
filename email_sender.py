"""Sends generated MIS reports by email. SMTP infrastructure (host/port,
sender account, password) is read from .env - no addresses or credentials
are hardcoded here. Per-client To/CC recipients come from the mapping file
the user uploads (see recipient_mapping.py), not from .env - a report is
only emailed to a client whose client_external_id appears in that mapping.
"""
import os
import smtplib
from email.message import EmailMessage

from dotenv import load_dotenv

load_dotenv()

REQUIRED_SMTP_VARS = ["SMTP_HOST", "SMTP_PORT", "EMAIL_SENDER", "EMAIL_PASSWORD"]


class EmailConfigError(RuntimeError):
    """Raised when required .env SMTP configuration is missing, or no
    recipient was provided for a report."""


def _get_smtp_config():
    missing = [name for name in REQUIRED_SMTP_VARS if not os.environ.get(name)]
    if missing:
        raise EmailConfigError(
            f"Missing required SMTP configuration in .env: {', '.join(missing)}"
        )
    return {
        "host": os.environ["SMTP_HOST"],
        "port": int(os.environ["SMTP_PORT"]),
        "sender": os.environ["EMAIL_SENDER"],
        "password": os.environ["EMAIL_PASSWORD"],
        "subject_template": os.environ.get(
            "EMAIL_SUBJECT_TEMPLATE", "Checkwise TAT Analysis Tracker {client_name}"
        ),
    }


def send_report_email(client_name, date_range_label, attachment_path, to_addresses, cc_addresses=None):
    """Sends the workbook at attachment_path to to_addresses (+ optional
    cc_addresses). Raises EmailConfigError if no To address was given, or
    EmailConfigError/smtplib errors on SMTP configuration/send failures -
    the caller decides how to surface those."""
    if not to_addresses:
        raise EmailConfigError(f'No To address available for {client_name} - not sent.')

    config = _get_smtp_config()
    cc_addresses = cc_addresses or []

    msg = EmailMessage()
    msg["Subject"] = config["subject_template"].format(client_name=client_name, date_range=date_range_label)
    msg["From"] = config["sender"]
    msg["To"] = ", ".join(to_addresses)
    if cc_addresses:
        msg["Cc"] = ", ".join(cc_addresses)
    msg.set_content(
        f"Dear {client_name},\n\n"
        "Thank you for using background verification services from AuthBridge.\n\n"
        "Please download consolidated Checkwise progress tracker:\n\n"
        "This tracker will give you a complete view of the current status of all cases received from "
        "your end according to the cases and checks requested.\n\n"
        "Please feel free to reach out to your SPOC for any clarifications.\n\n"
        "We are constantly striving to deliver excellent services by leveraging technology and making "
        "our operations more efficient and quick.\n\n"
        "Thank you for your time and business and helping us in shaping the future of background "
        "screening. We look forward to your continuous collaboration.\n\n\n\n"
        "Best Regards,\n"
        "Team AuthBridge\n"
        "www.authbridge.com"
    )

    with open(attachment_path, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=os.path.basename(attachment_path),
        )

    all_recipients = to_addresses + cc_addresses
    with smtplib.SMTP(config["host"], config["port"], timeout=120) as server:
        server.starttls()
        server.login(config["sender"], config["password"])
        server.send_message(msg, from_addr=config["sender"], to_addrs=all_recipients)
