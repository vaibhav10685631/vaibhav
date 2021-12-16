import requests
import datetime

MICROSOFT_OAUTH2_URL = "https://login.microsoftonline.com"
MICROSOFT_OAUTH2_PATH_BOT = "botframework.com/oauth2/v2.0/token"
MICROSOFT_OAUTH2_PATH_APP = "organizations/oauth2/v2.0/token"

bot_headers = None
token_expiration_date = datetime.datetime.now()

def refresh_token():
    clientID = "2e0d6109-4251-4003-b8ff-1d0ee6faf010"
    secretID = "aaduElqRs1nWro1_5_M-1dze_a-by-624j"
    redirectURI = "http://localhost:10060/oauth"
    f = open("refresh_token.txt", "r")
    rf_token = f.read()
    f.close()
    url = f"{MICROSOFT_OAUTH2_URL}/{MICROSOFT_OAUTH2_PATH_APP}"
    headers = {'content-type':'application/x-www-form-urlencoded'}
    payload = ("grant_type=refresh_token&scope=offline_access+User.ReadBasic.All+Chat.Create+ChatMember.ReadWrite+TeamsAppInstallation.ReadWriteForChat+AppCatalog.Read.All&client_id={0}&client_secret={1}&"
                        "refresh_token={2}&redirect_uri={3}").format(clientID, secretID, rf_token, redirectURI)

    response = requests.post(url=url, data=payload, headers=headers)

    if response.ok:
        data = response.json()
        access_token = data["access_token"]
        f = open("access_token.txt","w")
        f.write(access_token)
        f.close()

        headers = {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + access_token
        }
        return headers
    else:
        print("Could not get Access Token")

    return headers

def get_bot_headers():
    global token_expiration_date, bot_headers
    clientID = "2d8a2219-185f-4e3f-be55-5d335e766b50"
    secretID = "AtjfpSX4R5h.N.~BK1C8x4YThoQ-5lCQid"
    if token_expiration_date < datetime.datetime.now():
        uri = f"{MICROSOFT_OAUTH2_URL}/{MICROSOFT_OAUTH2_PATH_BOT}"
        grant_type = "client_credentials"
        scope = "https://api.botframework.com/.default"
        payload = {
            "client_id": clientID,
            "client_secret": secretID,
            "grant_type": grant_type,
            "scope": scope,
        }

        token_response = requests.post(uri, data=payload)

        if token_response.ok:
            token_data = token_response.json()
            access_token = token_data["access_token"]
            token_expiration = token_data["expires_in"]
            
            delta = datetime.timedelta(seconds=int(token_expiration))
            token_expiration_date = datetime.datetime.now() + delta

            bot_headers = {
                "content-type": "application/json",
                "Authorization": "Bearer %s" % access_token,
            }
            return bot_headers
        else:
            print("Could not get BotFramework token")
    else:
        return bot_headers