import requests
import json
from actions.auth_tokens import get_bot_headers

UPDATE_URL =  "https://smba.trafficmanager.net/in/v3/conversations/{0}/activities/{1}"

def update_activity(conversationId: str, activityId: str, user: str, card_content = "The ticket has been updated!"):
    url = UPDATE_URL.format(conversationId,activityId)
    headers = get_bot_headers()
        
    updated_card = {
        "type": "message",
        "attachments":[
            {
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content":
                {
                    "type": "AdaptiveCard",
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "version": "1.2",
                    "body": [
                        {
                            "type": "TextBlock",
                            "size": "medium",
                            "text": card_content,
                            "color": "good",
                            "weight": "bolder",
                            "wrap": True
                        },
                        {
                            "type": "FactSet",
                            "facts": [
                                {
                                    "title": "Executed By",
                                    "value": user
                                }
                            ]
                        }
                    ]
                }
            }
        ]
    }

    response = requests.put(url, headers=headers, data=json.dumps(updated_card))

    if not response.ok:
        print("Error trying to update card. Response: %s",response.text)

    return