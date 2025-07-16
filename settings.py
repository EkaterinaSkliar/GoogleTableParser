import httplib2
import apiclient
from oauth2client.service_account import (
    ServiceAccountCredentials,
)

LIST_SPREADSHEET_IDS = [
    '14DAHBO6FUe4KbWsMKaN9vhAqE_IBuXqcssBoQIEpLls',
    '130NJRk4iDHKJrmedSz2mRiLmL1k2RERB-5-yfxO_6jI',
]

GENERAL_STREADSHEET_ID = '1MSGBo3Mk3lqahaQd4VapmkZ86VVIOde1ytN0hAEVOa8'

CREDENTIALS_FILE = 'creds.json'

credentials = ServiceAccountCredentials.from_json_keyfile_name(
    CREDENTIALS_FILE,
    [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive',
    ],
)
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build(
    serviceName='sheets',
    version='v4',
    http=httpAuth,
)

TABLE_HEADER = 3
