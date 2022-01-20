"""
This file contains the global variables.
"""
import datetime

with open("access_token.txt", "r", encoding='utf8') as token_file:
    ACCESS_TOKEN = token_file.read()

HEADERS = {
    'Content-Type': 'application/json',
    'Authorization': 'Bearer ' + ACCESS_TOKEN
}

BOT_HEADERS = None
TOKEN_EXPIRATION_DATE = datetime.datetime.now()
