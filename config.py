"""
Configurations for MS Exchange
"""

ALLOWED_EXTENSIONS = "pdf, doc, docx, xls, xlsm, xlsx, ppt, pptx"
DOWNLOAD_ATTACHED_EMAILS = False
DOWNLOAD_ROOT_PATH = "/tmp/ms-attachments"
OUTGOING_SERVER = "outlook.office365.com"
RANGE_IN_SECONDS = 60 * 60 * 24 * 90  # past 90 days
TIMEZONE = "Asia/Singapore"