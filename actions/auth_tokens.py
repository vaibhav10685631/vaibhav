"""
This file contains the credentials and functions for Authentication.
"""

import datetime
import logging
from sre_constants import SUCCESS
import requests
import actions.globals

logger = logging.getLogger(__name__)

MICROSOFT_OAUTH2_URL = "https://login.microsoftonline.com"
MICROSOFT_OAUTH2_PATH_BOT = "botframework.com/oauth2/v2.0/token"
MICROSOFT_OAUTH2_PATH_APP = "organizations/oauth2/v2.0/token"

APP_CLIENT_ID = "2e0d6109-4251-4003-b8ff-1d0ee6faf010"
APP_SECRET_ID = "u4eDKJMlRjx5uJsdeA04ENV1NEJQfGkOoo"
SCOPE = "offline_access+User.ReadBasic.All+Chat.Create+ChatMember.ReadWrite+"\
    "TeamsAppInstallation.ReadWriteForChat+AppCatalog.Read.All"
REDIRECT_URI = "http://localhost:10060/oauth"

BOT_CLIENT_ID = "2d8a2219-185f-4e3f-be55-5d335e766b50"
BOT_SECRET_ID = "AtjfpSX4R5h.N.~BK1C8x4YThoQ-5lCQid"

def refresh_token():
    """Refreshes the Auth Token for App"""
    with open("refresh_token.txt", "r", encoding='utf8') as token_file:
        RF_TOKEN = token_file.read()

    url = f"{MICROSOFT_OAUTH2_URL}/{MICROSOFT_OAUTH2_PATH_APP}"
    headers = {'content-type':'application/x-www-form-urlencoded'}
    payload = f"grant_type=refresh_token&scope={SCOPE}&client_id={APP_CLIENT_ID}"\
        f"&client_secret={APP_SECRET_ID}&refresh_token={RF_TOKEN}&redirect_uri={REDIRECT_URI}"

    response = requests.post(url=url, data=payload, headers=headers)

    if response.ok:
        data = response.json()
        access_token = data["access_token"]
        with open("access_token.txt","w", encoding='utf8') as file:
            file.write(access_token)

        with open("refresh_token.txt","w", encoding='utf8') as file:
            file.write(data["refresh_token"])

        headers = {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + access_token
        }
        actions.globals.HEADERS = headers
        return "success"
    else:
        logger.error("Could not get Access Token")
        return "fail"

def get_bot_headers():
    """Returns the headers with BOT Token"""

    token_expiration_date = actions.globals.TOKEN_EXPIRATION_DATE

    if token_expiration_date < datetime.datetime.now():
        uri = f"{MICROSOFT_OAUTH2_URL}/{MICROSOFT_OAUTH2_PATH_BOT}"
        grant_type = "client_credentials"
        scope = "https://api.botframework.com/.default"
        payload = {
            "client_id": BOT_CLIENT_ID,
            "client_secret": BOT_SECRET_ID,
            "grant_type": grant_type,
            "scope": scope,
        }

        token_response = requests.post(uri, data=payload)

        if token_response.ok:
            token_data = token_response.json()
            access_token = token_data["access_token"]
            token_expiration = token_data["expires_in"]

            delta = datetime.timedelta(seconds=int(token_expiration))
            actions.globals.TOKEN_EXPIRATION_DATE = datetime.datetime.now() + delta

            actions.globals.BOT_HEADERS = {
                "content-type": "application/json",
                "Authorization": f"Bearer {access_token}"
            }
        else:
            logger.error("Could not get BotFramework token")
            return None

    return actions.globals.BOT_HEADERS
