import os

import gspread
from oauth2client.service_account import ServiceAccountCredentials

CREDS_JSON = os.environ['CREDS_JSON']


def get_client(creds_json=CREDS_JSON):
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(creds_json, scope)
    client = gspread.authorize(creds)

    return client


def open_workbook(wb_name, client=None):
    if not client:
        client = get_client()

    open_wb = client.open(wb_name)

    return wb_name, open_wb
