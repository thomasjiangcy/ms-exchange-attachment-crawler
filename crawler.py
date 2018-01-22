"""
Simple script to crawl through user's MS Exchange inbox to download attachments from emails
"""

import os
import errno
from datetime import datetime, timedelta
from getpass import getpass

import pytz
from exchangelib import (
    DELEGATE, Account, Configuration, EWSDateTime, FileAttachment,
    ItemAttachment, Message, ServiceAccount
)

from config import (
    ALLOWED_EXTENSIONS, DOWNLOAD_ATTACHED_EMAILS,
    DOWNLOAD_ROOT_PATH, OUTGOING_SERVER, RANGE_IN_SECONDS,
    TIMEZONE
)


PARSED_EXTENSIONS = [ext for ext in (x.strip() for x in ALLOWED_EXTENSIONS.split(","))]


def get_user_login():
    """Get user login credentials"""
    # User login
    username, password = ("", "")

    while username == "":
        username = input("Enter username: ")

    email = input("Enter email (optional if username is email): ")

    while password == "":
        password = getpass("Enter password: ")

    return (username.strip(), email.strip(), password)

def login(username, email, password):
    """Login to MS Exchange account"""
    print("Logging in...")
    # Construct login credentials with fault tolerant ServiceAccount
    credentials = ServiceAccount(username=username, password=password)
    config = Configuration(server=OUTGOING_SERVER, credentials=credentials)

    # Retrieve account
    account = Account(
        primary_smtp_address=email if bool(email) else username,
        autodiscover=False,
        config=config,
        access_type=DELEGATE
    )

    print("Login successful.")

    return account

def check_directories(path):
    """Checks if directories exist along path and create accordingly"""
    if not os.path.exists(os.path.dirname(path)):
        try:
            os.makedirs(os.path.dirname(path))
        except OSError as exc: # Guard against race condition
            if exc.errno != errno.EEXIST:
                raise

def is_valid_extension(filename):
    """Checks if file is of correct extension"""
    ext = filename.split(".")[-1]
    return ext in PARSED_EXTENSIONS

def get_attachments(inbox):
    """Downloads all attachments to host machine"""
    start_date = datetime.now() - timedelta(seconds=RANGE_IN_SECONDS)
    year, month, date = (start_date.year, start_date.month, start_date.day)
    ews_start_date = pytz.timezone(TIMEZONE).localize(EWSDateTime(year, month, date + 1))

    print("Retrieving attachments from {0}...".format(ews_start_date))
    qs = inbox.filter(datetime_received__gte=ews_start_date)

    for item in inbox.all():
        formatted_datetime = datetime.strftime(item.datetime_received, "%Y-%m-%d-%H-%M-%S")

        for attachment in item.attachments:
            if isinstance(attachment, FileAttachment) and is_valid_extension(attachment.name):
                local_path = os.path.join(DOWNLOAD_ROOT_PATH, formatted_datetime, attachment.name)
                check_directories(local_path)
                with open(local_path, 'wb') as f:
                    f.write(attachment.content)
                print("Saved attachment to {0}".format(local_path))
            
            elif isinstance(attachment, ItemAttachment) and DOWNLOAD_ATTACHED_EMAILS:
                if isinstance(attachment.item, Message):
                    local_path = os.path.join(DOWNLOAD_ROOT_PATH, formatted_datetime, attachment.item.subject)
                    check_directories(local_path)
                    with open(local_path, 'wb') as f:
                        f.write(attachment.item.body)
                    print("Saved email to {0}".format(local_path))
            
            else:
                print("Skipping..")

def run():
    """Executes script"""
    username, email, password = get_user_login()
    account = login(username, email, password)
    inbox = account.inbox
    get_attachments(inbox)

if __name__ == "__main__":
    run()
