"""
This files contains all the custom actions which can be used to run IIM funtions.
"""

import json
from typing import Any, Text, Dict, List
import datetime
import requests
from mailmerge import MailMerge

from rasa_sdk import Action, Tracker
from rasa_sdk.executor import CollectingDispatcher
from rasa_sdk.events import ReminderScheduled, ReminderCancelled, SlotSet

from actions.requester import ENGINE, get_response,get_article,put_response,get_file_content
from actions.requester import post_attachment,get_attachment,get_groups
from actions.auth_tokens import refresh_token, get_bot_headers
from actions.card_activity import update_activity
from actions.email_notification import send_email_notification, send_mir, get_recipients

from actions.constants import SLA_TABLE_SPEC, INC_TABLE_SPEC, JRNL_TABLE_SPEC, KNLDG_TABLE_SPEC
from actions.constants import EMAIL_SLOTS, BOT_URL, CATALOG_APP_ID
from actions.constants import TENANT, SITE_ID, SITE_NAME, FOLDER
import actions.globals

######### IIM Custom Actions #########

########### Incident Trigger from ITSM ###########
class ActionNewIncident(Action):
    """Performs actions when new major incident is created"""

    def name(self) -> Text:
        return "action_new_incident"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        ######## Get incident info from trigger intent entities ########
        sys_id = tracker.get_slot('sys_id')
        number = tracker.get_slot('number')
        short_description = tracker.get_slot('short_description')
        priority = tracker.get_slot('Epriority')
        caller = tracker.get_slot("caller")

        business_impact = tracker.get_slot('EBizImp')
        if not business_impact:
            business_impact = "(Impact Assessment is in progress.)"

        sys_created_on = tracker.get_slot('sys_created_on')
        inc_start = datetime.datetime.strptime(sys_created_on, '%d-%m-%Y %H:%M:%S').strftime("%d-%b-%Y %H:%M")

        ########### 0. Check if the room/group already exists for the particular incident ###################
        query = f"SELECT * FROM incident_chat_map WHERE number='{number}'"
        result = ENGINE.execute(query)
        row = result.fetchall()
        print("\n Result row ::", row)
        if len(row) == 1:
            chat_id = row[0]['chat_id']
            bot_conv_url =  BOT_URL+chat_id+"/activities"
            bot_headers = get_bot_headers()
            inc_update = {
                "type":"message",
                "text" : "<b>The incident is again promoted to Major Incident State!</b>"
            }

            send_update_response = requests.post(bot_conv_url, headers=bot_headers, data=json.dumps(inc_update))

            if not send_update_response.ok:
                print("Error trying to send ticket update message. Response: %s",send_update_response.text)

        else:
            ########### 1. Create a New Chat and add Members in Teams ###################

            url = "https://graph.microsoft.com/v1.0/chats"

            member_ids = ['bhakti.prabhu@iimbot.onmicrosoft.com', 'amol.chaudhari@iimbot.onmicrosoft.com', 'karanjeet.singh@iimbot.onmicrosoft.com']
            members = []
            for member_id in member_ids:
                members.append(
                    {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["owner"],
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{member_id}')"
                    }
                )

            data = {
                "chatType": "group",
                "topic": "MIM "+number,
                "members": members
            }

            data = json.dumps(data)
            while True:
                response = requests.post(url, headers=actions.globals.HEADERS, data=data)
                if response.status_code == 401:
                    print("Token Expired.....Refreshing Token")
                    refresh_token()
                elif response.ok:
                    chat_data = response.json()
                    chat_id = chat_data['id']
                    print("*****Chat created successfully*****")
                    break
                else:
                    print(
                        'Status:', response.status_code,
                        'Headers:', response.headers,
                        'Error Response:',response.json()
                    )
                    return []

            ######## 2. Map Number and Room Id and Store it in the Database #########
            query = f"INSERT INTO incident_chat_map VALUES ('{number}','{sys_id}','{chat_id}')"
            ENGINE.execute(query)

            ########## 3. Add Bot to the Created Chat ##########
            url = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/installedApps"
            data = {
            "teamsApp@odata.bind":"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/"+CATALOG_APP_ID
            }

            data = json.dumps(data)

            while True:
                response = requests.post(url, headers=actions.globals.HEADERS, data=data)
                if response.status_code == 401:
                    print("Token Expired.....Refreshing Token")
                    refresh_token()
                elif response.ok:
                    print("*****App Installed Successfully*****")
                    break
                else:
                    print(
                        'Status:', response.status_code,
                        'Headers:', response.headers,
                        'Error Response:',response.json()
                    )
                    break

            ###### 3. Alert Message By Bot #######
            bot_conv_url =  BOT_URL+chat_id+"/activities"
            bot_headers = get_bot_headers()
            new_ticket_alert = {
                "type":"message",
                "attachments": [
                    {
                        "contentType": "application/vnd.microsoft.card.adaptive",
                        "content":
                        {
                            "type": "AdaptiveCard",
                            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                            "version": "1.2",
                            "body": [
                                {
                                    "type": "ColumnSet",
                                    "columns": [
                                        {
                                            "type": "Column",
                                            "items": [
                                                {
                                                    "type": "TextBlock",
                                                    "weight": "Bolder",
                                                    "text": "A Major Incident is raised!",
                                                    "color": "Warning",
                                                    "size": "Large",
                                                    "spacing": "Small"
                                                }
                                            ],
                                            "width": "stretch"
                                        }
                                    ]
                                },
                                {
                                    "type": "ColumnSet",
                                    "columns": [
                                        {
                                            "type": "Column",
                                            "width": 35,
                                            "items": [
                                                {
                                                    "type": "TextBlock",
                                                    "text": "Incident Number:"
                                                },
                                                {
                                                    "type": "TextBlock",
                                                    "text": "Created On:",
                                                    "spacing": "Small"
                                                },
                                                {
                                                    "type": "TextBlock",
                                                    "text": "Description:",
                                                    "spacing": "Small"
                                                }
                                            ]
                                        },
                                        {
                                            "type": "Column",
                                            "width": 65,
                                            "items": [
                                                {
                                                    "type": "TextBlock",
                                                    "text": number
                                                },
                                                {
                                                    "type": "TextBlock",
                                                    "text": inc_start,
                                                    "spacing": "Small"
                                                },
                                                {
                                                    "type": "TextBlock",
                                                    "text": short_description,
                                                    "spacing": "Small"
                                                }
                                            ]
                                        }
                                    ],
                                    "spacing": "Padding",
                                    "horizontalAlignment": "Center"
                                }
                            ],
                        }
                    }
                ]
            }

            send_alert_response = requests.post(bot_conv_url, headers=bot_headers, data=json.dumps(new_ticket_alert))

            if not send_alert_response.ok:
                print("Error trying to send new ticket alert message. Response: %s",send_alert_response.text)

            ###### 4. Send Email Details Card #######
            next_update = datetime.datetime.strptime(inc_start,'%d-%b-%Y %H:%M') + datetime.timedelta(minutes=30)
            mim_eng = datetime.datetime.strptime(inc_start,'%d-%b-%Y %H:%M') + datetime.timedelta(minutes=10)
            nxt_upd_due = next_update.strftime("%d-%b-%Y %H:%M")
            mim_eng_time = mim_eng.strftime("%d-%b-%Y %H:%M")

            recipients = ",".join(get_recipients())

            email_details_card = {
                "type":"message",
                "attachments": [
                    {
                        "contentType": "application/vnd.microsoft.card.adaptive",
                        "content":
                        {
                            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                            "type": "AdaptiveCard",
                            "version": "1.2",
                            "body": [
                                {
                                    "type": "TextBlock",
                                    "size": "medium",
                                    "weight": "bolder",
                                    "text": "Email Communication Details",
                                    "horizontalAlignment": "center"
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Priority:",
                                            "color": "Accent",
                                            "weight": "Bolder"
                                        }
                                    ]
                                },
                                {
                                    "type": "Input.ChoiceSet",
                                    "choices": [
                                        {
                                            "title": "P1",
                                            "value": "P1"
                                        },
                                        {
                                            "title": "P2",
                                            "value": "P2"
                                        }
                                    ],
                                    "id": "Epriority",
                                    "value": priority,
                                    "spacing": "None"
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Major Incident State:",
                                            "color": "Accent",
                                            "weight": "Bolder"
                                        }
                                    ]
                                },
                                {
                                    "type": "Input.ChoiceSet",
                                    "choices": [
                                        {
                                            "title": "Declared",
                                            "value": "Declared"
                                        },
                                        {
                                            "title": "In-Progress",
                                            "value": "In-Progress"
                                        },
                                        {
                                            "title": "On-Hold",
                                            "value": "On-Hold"
                                        },
                                        {
                                            "title": "Under Observation",
                                            "value": "Under Observation"
                                        },
                                        {
                                            "title": "Restored/Monitoring",
                                            "value": "Restored/Monitoring"
                                        },
                                        {
                                            "title": "Resolved",
                                            "value": "Resolved"
                                        },
                                        {
                                            "title": "Downgraded",
                                            "value": "Downgraded"
                                        },
                                        {
                                            "title": "Cancelled",
                                            "value": "Cancelled"
                                        }
                                    ],
                                    "id": "Emistate",
                                    "value": "Declared",
                                    "spacing": "None"
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Incident Summary:",
                                            "color": "Accent",
                                            "weight": "Bolder"
                                        }
                                    ]
                                },
                                {
                                    "type": "Input.Text",
                                    "id": "EIncSummary",
                                    "spacing": "None",
                                    "value": short_description
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Business Impact (Description):",
                                            "color": "Accent",
                                            "weight": "Bolder"
                                        }
                                    ]
                                },
                                {
                                    "type": "Input.Text",
                                    "id": "EBizImp",
                                    "spacing": "None",
                                    "value": business_impact
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Impacted Location(s) / Sites:",
                                            "color": "Accent",
                                            "weight": "Bolder"
                                        }
                                    ]
                                },
                                {
                                    "type": "Input.Text",
                                    "id": "EImpLoc",
                                    "spacing": "None",
                                    "value": ""
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Impacted Clients:",
                                            "color": "Accent",
                                            "weight": "Bolder"
                                        }
                                    ]
                                },
                                {
                                    "type": "Input.Text",
                                    "id": "EImpClient",
                                    "spacing": "None",
                                    "value": ""
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Impacted Application / Services:",
                                            "color": "Accent",
                                            "weight": "Bolder"
                                        }
                                    ]
                                },
                                {
                                    "type": "Input.Text",
                                    "id": "EImpApp",
                                    "spacing": "None",
                                    "value": ""
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "No. of users impacted:",
                                            "color": "Accent",
                                            "weight": "Bolder"
                                        }
                                    ]
                                },
                                {
                                    "type": "Input.Text",
                                    "id": "EImpUsr",
                                    "spacing": "None",
                                    "value": ""
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Issue reported by:",
                                            "color": "Accent",
                                            "weight": "Bolder"
                                        }
                                    ],
                                    "separator": True
                                },
                                {
                                    "type": "Input.Text",
                                    "id": "EIsRepBy",
                                    "spacing": "None",
                                    "value": caller
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Incident Ticket Ref#",
                                            "color": "Accent",
                                            "weight": "Bolder"
                                        }
                                    ]
                                },
                                {
                                    "type": "Input.Text",
                                    "id": "EIncRef",
                                    "spacing": "None",
                                    "value": number
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Vendor Name / Ticket Ref:",
                                            "color": "Accent",
                                            "weight": "Bolder"
                                        }
                                    ]
                                },
                                {
                                    "type": "Input.Text",
                                    "id": "EVendor",
                                    "spacing": "None",
                                    "value": "NA"
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Incident Start Date/ Time:",
                                            "color": "Accent",
                                            "weight": "Bolder"
                                        }
                                    ]
                                },
                                {
                                    "type": "Input.Text",
                                    "id": "EIncStart",
                                    "spacing": "None",
                                    "value": inc_start
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Date / Time (MIM Engaged):",
                                            "color": "Accent",
                                            "weight": "Bolder"
                                        }
                                    ]
                                },
                                {
                                    "type": "Input.Text",
                                    "id": "EMimEng",
                                    "spacing": "None",
                                    "value": mim_eng_time
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Major Incident Manager:",
                                            "color": "Accent",
                                            "weight": "Bolder"
                                        }
                                    ],
                                    "separator": True
                                },
                                {
                                    "type": "Input.ChoiceSet",
                                    "choices": [
                                        {
                                            "title": "MIM Mgr 1",
                                            "value": "MIM Mgr 1"
                                        },
                                        {
                                            "title": "MIM Mgr 2",
                                            "value": "MIM Mgr 2"
                                        },
                                        {
                                            "title": "MIM Mgr 3",
                                            "value": "MIM Mgr 3"
                                        }
                                    ],
                                    "id": "EMIM",
                                    "value": "",
                                    "spacing": "None"
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Support Teams involved:",
                                            "color": "Accent",
                                            "weight": "Bolder"
                                        }
                                    ]
                                },
                                {
                                    "type": "Input.Text",
                                    "id": "ESupTeams",
                                    "spacing": "None",
                                    "value": ""
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Workaround:",
                                            "color": "Accent",
                                            "weight": "Bolder"
                                        }
                                    ],
                                    "separator": True
                                },
                                {
                                    "type": "Input.Text",
                                    "id": "EWrkArnd",
                                    "spacing": "None",
                                    "value": "(To be determind)"
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Change Related / Ref:",
                                            "color": "Accent",
                                            "weight": "Bolder"
                                        }
                                    ],
                                    "separator": True
                                },
                                {
                                    "type": "Input.Text",
                                    "id": "EChange",
                                    "spacing": "None",
                                    "value": "(To be determind)"
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Reason for Outage (RFO):",
                                            "color": "Accent",
                                            "weight": "Bolder"
                                        }
                                    ]
                                },
                                {
                                    "type": "Input.Text",
                                    "id": "ERFO",
                                    "spacing": "None",
                                    "value": "(To be determind)"
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Actual Resolution Time:",
                                            "color": "Accent",
                                            "weight": "Bolder"
                                        }
                                    ]
                                },
                                {
                                    "type": "Input.Text",
                                    "id": "EResTime",
                                    "spacing": "None",
                                    "value": "(To be determind)"
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Outage Duration:",
                                            "color": "Accent",
                                            "weight": "Bolder"
                                        }
                                    ]
                                },
                                {
                                    "type": "Input.Text",
                                    "id": "EOutDur",
                                    "spacing": "None",
                                    "value": "(To be determind)"
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Next Update Due:",
                                            "color": "Accent",
                                            "weight": "Bolder"
                                        }
                                    ]
                                },
                                {
                                    "type": "Input.Text",
                                    "id": "ENxtUpd",
                                    "spacing": "None",
                                    "value": nxt_upd_due
                                },
                                {
                                    "type": "Container",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Distribution Lists: ",
                                            "color": "Accent",
                                            "weight": "Bolder"
                                        }
                                    ],
                                    "separator": True
                                },
                                {
                                    "type": "Input.Text",
                                    "id": "EDL",
                                    "spacing": "None",
                                    "value": recipients,
                                    "isMultiline": True
                                },
                                {
                                    "type": "Input.Toggle",
                                    "title": "Recieve Reminders for sending Email Notification",
                                    "id": "Erem_flag"
                                },
                                {
                                    "type": "ActionSet",
                                    "actions": [
                                        {
                                            "type": "Action.Submit",
                                            "title": "Save Details",
                                            "data": {
                                                "msteams": {
                                                    "type": "messageBack",
                                                    "text": "User interaction with Email Details card"
                                                },
                                                "init_com": "true"
                                            }
                                        }
                                    ],
                                    "horizontalAlignment": "Center",
                                    "spacing": "None"
                                }
                            ]
                        }
                    }
                ]
            }

            email_details_response = requests.post(bot_conv_url, headers=bot_headers, data=json.dumps(email_details_card))

            if not send_alert_response.ok:
                print("Error trying to send email details card. Response: %s",email_details_response.text)

        ######## 7. Store Room Id in Snow Incident Table #########
        query_filter = '?sysparm_fields=u_room_id'
        data = {
            "u_room_id": chat_id
        }
        response_status = put_response(INC_TABLE_SPEC, chat_id, query_filter, json.dumps(data))

        if response_status is not None:
            print("\n ******Room Id Stored in the SNOW Incident Table******")

        return []

class ActionIncidentUpdate(Action):
    """Performs actions when a major incident is updated"""

    def name(self) -> Text:
        return "action_incident_update"

    async def run(self, dispatcher: CollectingDispatcher,
            tracker: Tracker,
            domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        first_update = tracker.get_slot('fst_update')
        resolved_flag = tracker.get_slot("res_flag")

        number = tracker.get_slot('number')
        state = tracker.get_slot('state')
        mi_state = tracker.get_slot('mi_state')
        note_info = tracker.get_slot('note_info')
        note_msg = tracker.get_slot('note_msg')
        note_type = tracker.get_slot('note_type')

        updated_by = note_info.split('-')[-1]
        chat_id = tracker.sender_id

        bot_conv_url =  BOT_URL+chat_id+"/activities"
        bot_headers = get_bot_headers()

        if mi_state == 'Canceled':
            print("\n\nThe incident in demoted from major incident state.")
            inc_update = {
                "type": "message",
                "text" : "**This incident is no longer a Major Incident.**<br>"\
                    "*Hence, the Update Notifications and Reminders for this incident will be stopped.*<br>"\
                    "Update Information: - <br> **Updated By:** "\
                    +updated_by+"<br> **"+note_type+":** "+note_msg.replace('\n','<br>')
                }

            send_update_response = requests.post(bot_conv_url, headers=bot_headers, data=json.dumps(inc_update))

            if not send_update_response.ok:
                print("Error trying to send ticket update message. Response: %s",send_update_response.text)

            ##### Cancel All Reminders #####
            no_update_rem = "no_update_"+number
            print("\nCancelling Reminders..............")
            print("\n", no_update_rem, "\n email_update_rem ")

            return [ReminderCancelled(name=no_update_rem), ReminderCancelled("email_update_rem")]

        if state in ('Resolved','Closed'):
            print("The incident is Resolved/Closed")
            response = "**The incident has been "+state+".**"
            if state == "Resolved" and resolved_flag == 'false':
                response = response + "<br> Thank you all for your participation and contribution."\
                    " <br> Have a good day :)"

                ###### Adding Resolution Summary to response #######
                query_filter = "?sysparm_display_value=True&sysparm_exclude_reference_link=True"\
                    "&sysparm_fields=number,resolved_at,resolved_by,knowledge,close_code,close_notes"
                inc_data = get_response(INC_TABLE_SPEC, chat_id, query_filter)
                if inc_data is not None:
                    resolved_at = inc_data['resolved_at']
                    resolved_by = inc_data['resolved_by']
                    knowledge = inc_data['knowledge']
                    resolution_code = inc_data['close_code']
                    resolution_notes = inc_data['close_notes']

                    query_filter = "^sla.targetINresolution^sla.typeINSLA&sysparm_display_value=True"\
                        "&sysparm_exclude_reference_link=True&sysparm_fields=has_breached,business_duration"
                    sla_data = get_response(SLA_TABLE_SPEC, chat_id, query_filter)
                    if sla_data is not None:
                        has_breached = sla_data[0]['has_breached']
                        business_duration = sla_data[0]['business_duration']
                        if has_breached:
                            sla_met = 'Yes'
                        else:
                            sla_met = 'No'

                        if knowledge=="true":
                            knowledge_article = "\n\n *This resolution has been submitted to be published as knowledge article.*"
                        else:
                            knowledge_article = ""

                        response = response + "<br><br>Following is the summary of resolution of the incident: - "\
                            "<br> **Resolved At:** "+resolved_at+"<br> **Resolved By:** "+resolved_by+\
                            "<br> **SLA met:** "+sla_met+"<br> **Resolution Duration:** "+business_duration+\
                            "<br> **Resolution Code:** "+resolution_code+"<br> **Resolution Notes:** "\
                            +resolution_notes+knowledge_article

            elif resolved_flag == "true":
                print("\n\nThe incident information is updated")
                inc_update = {
                "type": "message",
                "text" : "There is a new update on the incident: - <br> **Updated By:** "\
                    +updated_by+"<br> **"+note_type+":** "+note_msg.replace('\n','<br>')
                }

                send_update_response = requests.post(bot_conv_url, headers=bot_headers, data=json.dumps(inc_update))

                if not send_update_response.ok:
                    print("Error trying to send ticket update message. Response: %s",send_update_response.text)

                return []

            inc_update = {
                "type": "message",
                "text" : response
            }

            send_update_response = requests.post(bot_conv_url, headers=bot_headers, data=json.dumps(inc_update))

            if not send_update_response.ok:
                print("Error trying to send botframework message. Response: %s",send_update_response.text)

            ###### Card for final email update ######

            # final_email_card = []

            ##### Cancel All Reminders #####
            no_update_rem = "no_update_"+number
            print("\nCancelling Reminders..............")
            print("\n", no_update_rem, "\n email_update_rem ")

            return [ReminderCancelled(name=no_update_rem), ReminderCancelled(name="email_update_rem"), SlotSet('res_flag','true')]

        events_list = []
        if resolved_flag == "true":
            events_list.extend([SlotSet('res_flag','false')])

        if first_update == "false":
            print("\n\nThe incident information is updated")
            inc_update = {
                "type":"message",
                "text" : "There is a new update on the incident: - <br> **Updated By:** "\
                    +updated_by+"<br> **"+note_type+":** "+note_msg.replace('\n','<br>')
            }

            send_update_response = requests.post(bot_conv_url, headers=bot_headers, data=json.dumps(inc_update))

            if not send_update_response.ok:
                print("Error trying to send ticket update message. Response: %s",send_update_response.text)
        else:
            events_list.extend([SlotSet('fst_update','false')])

        ####### Re-setting Reminder for No Update #######
        date = datetime.datetime.now() + datetime.timedelta(minutes=5)
        name = "no_update_"+number
        no_update_reminder = ReminderScheduled(
            "EXTERNAL_no_update_reminder",
            trigger_date_time=date,
            entities={"chat_id": chat_id, "number": number, "no_upd_interval": 15},
            name=name,
            kill_on_user_message=False,
        )
        print("\n\nNo Update DateTime :: ",date)
        print("~~~~~~~~~~~~No Update Reminder is Set~~~~~~~~~~~~~~")

        events_list.extend([no_update_reminder])

        return events_list

############## Data Fetch actions ################
class ActionAssignedTo(Action):
    """Fetches Assigned Engineer from Incident Table in ITSM"""

    def name(self) -> Text:
        return "action_assigned_to"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        query_filter = '?sysparm_display_value=True&sysparm_exclude_reference_link=True&sysparm_fields=assigned_to'
        data = get_response(INC_TABLE_SPEC, tracker.sender_id, query_filter)
        if data is not None:
            assigned_to = data['assigned_to']
            if not assigned_to:
                dispatcher.utter_message(text = "The ticket is not assigned to anyone.")
            else:
                dispatcher.utter_message(text = "The ticket is assigned to **"+assigned_to+"**.")
        return []

class ActionAssignmentGroup(Action):
    """Fetches Assignment Group from Incident Table in ITSM"""

    def name(self) -> Text:
        return "action_assignment_group"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        query_filter = '?sysparm_display_value=True&sysparm_exclude_reference_link=True&sysparm_fields=assignment_group'
        data = get_response(INC_TABLE_SPEC, tracker.sender_id, query_filter)
        if data is not None:
            assignment_group = data['assignment_group']
            if not assignment_group:
                dispatcher.utter_message(text = "The ticket is not assigned to any group.")
            else:
                dispatcher.utter_message(text = "The ticket is assigned to **"+assignment_group+"** group.")
        return []

class ActionState(Action):
    """Fetches Incident State from Incident Table in ITSM"""

    def name(self) -> Text:
        return "action_state"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        query_filter = '?sysparm_display_value=True&sysparm_fields=state,hold_reason'
        data = get_response(INC_TABLE_SPEC, tracker.sender_id, query_filter)
        if data is not None:
            state = data['state']
            hold_reason = data['hold_reason']
            ans = "The state of the incident is: **"+state+"**."
            if state == "On Hold":
                ans = ans+"<br> On Hold Reason: "+hold_reason

            dispatcher.utter_message(text = ans)

        return []

class ActionReassignmentCount(Action):
    """Fetches Incident Reassignment Count from Incident Table in ITSM"""

    def name(self) -> Text:
        return "action_reassignment_count"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        query_filter = '?sysparm_fields=reassignment_count'
        data = get_response(INC_TABLE_SPEC, tracker.sender_id, query_filter)
        if data is not None:
            reassignment_count = data['reassignment_count']

            dispatcher.utter_message(text = "The reassignment count of the incident is **"+reassignment_count+"**.")

        return []

class ActionShortDescription(Action):
    """Fetches Short Description of Incident from Incident Table in ITSM"""

    def name(self) -> Text:
        return "action_short_description"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        query_filter = '?sysparm_fields=short_description'
        data = get_response(INC_TABLE_SPEC, tracker.sender_id, query_filter)
        if data is not None:
            short_description = data['short_description']

            dispatcher.utter_message(text = "The short description of the incident is:<br> *"+short_description+"*")

        return []

class ActionLongDescription(Action):
    """Fetches Long Description of Incident from Incident Table in ITSM"""

    def name(self) -> Text:
        return "action_long_description"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        query_filter = '?sysparm_fields=description'
        data = get_response(INC_TABLE_SPEC, tracker.sender_id, query_filter)
        if data is not None:
            long_description = data['description']
            if long_description == "":
                dispatcher.utter_message(text = "There is no long description for this incident.")
            else:
                dispatcher.utter_message(text = "The description of the incident is:<br> *"+long_description+"*")

            files = get_attachment(tracker.sender_id)
            if files:
                if len(files) != 0:
                    attachment_links = ''
                    for attachment in files:
                        attachment_links = attachment_links+' - ['+attachment['file_name']+']('+attachment['download_link']+')<br>'

                    reply = "Following are the files attached to the incident:<br>"\
                        +attachment_links+"<br>*(Click on the ones you wish to download.)*"

                    dispatcher.utter_message(text = reply)
        return []

class ActionTimeLeft(Action):
    """Fetches Time Left for the Incident to breach the SLA from Task SLA Table in ITSM"""

    def name(self) -> Text:
        return "action_time_left"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        query_filter = '^sla.targetINresolution^sla.typeINSLA&sysparm_display_value=True&sysparm_fields=time_left'
        data = get_response(SLA_TABLE_SPEC, tracker.sender_id, query_filter)
        if data is not None:
            time_left = data[0]['time_left']
            if int(time_left.split(' ')[0]) == 0:
                reply = "<span style='color: red;'>*The sla has been breached. There is no time left.*</span>"

            elif int(time_left.split(' ')[0]) == 1:
                reply = "There is **"+time_left+"** left for incident resolution."

            else:
                reply = "There are **"+time_left+"** left for incident resolution."

            dispatcher.utter_message(text = reply)

        return []

class ActionSlaDefinition(Action):
    """Fetches SLA Definition of the Incident from Task SLA Table in ITSM"""

    def name(self) -> Text:
        return "action_sla_definition"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        query_filter = "^sla.targetINresolution^sla.typeINSLA&sysparm_display_value=True"\
            "&sysparm_exclude_reference_link=True&sysparm_fields=sla"
        data = get_response(SLA_TABLE_SPEC, tracker.sender_id, query_filter)
        if data is not None:
            sla = data[0]['sla']

            reply = "The SLA defined for the incident is: **"+sla+"**"

            dispatcher.utter_message(text = reply)

        return []

class ActionElapsedPercentage(Action):
    """Fetches SLA Elapsed Percentage from Task SLA Table in ITSM"""

    def name(self) -> Text:
        return "action_elapsed_percentage"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        query_filter = '^sla.targetINresolution^sla.typeINSLA&sysparm_fields=percentage'
        data = get_response(SLA_TABLE_SPEC, tracker.sender_id, query_filter)
        if data is not None:
            percentage = data[0]['percentage']

            reply = "**"+percentage+"%** of SLA has been elapsed."

            dispatcher.utter_message(text = reply)

        return []

class ActionBreachTime(Action):
    """Fetches Breach Time of the Incident from Task SLA Table in ITSM"""

    def name(self) -> Text:
        return "action_breach_time"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        query_filter = '^sla.targetINresolution^sla.typeINSLA&sysparm_display_value=True&sysparm_fields=planned_end_time,time_left'
        data = get_response(SLA_TABLE_SPEC, tracker.sender_id, query_filter)
        if data is not None:
            breach_time = data[0]['planned_end_time']
            time_left = data[0]['time_left']

            reply = "The planned end time for this incident is: **"+breach_time+"**<br> Time Left: *"+time_left+"*"

            dispatcher.utter_message(text = reply)

        return []

class ActionLatestUpdate(Action):
    """Fetches Latest Worknote of the Incident from Journal Table in ITSM"""

    def name(self) -> Text:
        return "action_latest_update"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        query_filter = '^ORDERBYDESCsys_created_on&sysparm_fields=sys_created_on,sys_created_by,element,value&sysparm_limit=1'
        data = get_response(JRNL_TABLE_SPEC, tracker.sender_id, query_filter)
        if data is not None:
            updated_on = data[0]['sys_created_on']
            updated_by = data[0]['sys_created_by']
            element = data[0]['element']
            updated_value = data[0]['value']

            if element == 'work_notes':
                element = "Work Note"
            else:
                element = "Comment"

            reply = "Following are the details of the latest update on the incident: - <br> **Updated On:** "\
                +updated_on+"<br> **Updated By:** "+updated_by+"<br> **"+element+":** "+updated_value

            dispatcher.utter_message(text = reply)

        return []

class ActionKnowledgeArticle(Action):
    """Fetches Knowledge Article attached to the Incident from Knowledge Table in ITSM"""

    def name(self) -> Text:
        return "action_knowledge_article"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        query_filter1 = '&sysparm_exclude_reference_link=True&sysparm_fields=kb_knowledge'
        data = get_response(KNLDG_TABLE_SPEC, tracker.sender_id, query_filter1)
        if data is not None:
            if len(data) == 0:
                knowledge_articles = "*No, there aren't any knowledge articles attached to this incident.*"
            elif len(data) == 1:
                knowledge_articles = "Yes, there is 1 knowledge article attached to this incident. "\
                    "Following are the details of the article: -<br>"
            else:
                knowledge_articles = "Yes, there are "+len(data)+" knowledge articles attached to this incident. "\
                    "Following are the details of the articles: -<br>"

            for data_item in data:
                sys_id = data_item['kb_knowledge']
                article = get_article(sys_id)
                if article is not None:
                    number = article['number']
                    short_description = article['short_description']
                    view_count = article['sys_view_count']
                    use_count = article['use_count']
                    body = article['text']
                    knowledge_articles = knowledge_articles+"[**"+number+":** *"\
                        +short_description+"*](https://dev89325.service-now.com/kb_view.do?sys_kb_id="\
                        +sys_id+") <br> **View count:** "+view_count+"    **Use count:** "\
                        +use_count+"<br> *Article:*<br>"+body+"<br><br>"

            dispatcher.utter_message(text = knowledge_articles)

        return []

class ActionParent(Action):
    """Fetches Parent Incident of the current Incident from Incident Table in ITSM"""

    def name(self) -> Text:
        return "action_parent"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        query_filter = '?sysparm_display_value=True&sysparm_exclude_reference_link=True&sysparm_fields=parent_incident'
        data = get_response(INC_TABLE_SPEC, tracker.sender_id, query_filter)
        if data is not None:
            parent = data['parent_incident']
            if parent == "":
                reply = "No. This incident is not related to any other incident."
            else:
                reply = "Yes. This incident is a child incident of **"+parent+"**."

            dispatcher.utter_message(text = reply)

        return []

class ActionResolutionSummary(Action):
    """Fetches Resolution Information of the Incident from Incident Table in ITSM"""

    def name(self) -> Text:
        return "action_resolution_summary"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        query_filter = "?sysparm_display_value=True&sysparm_exclude_reference_link=True"\
            "&sysparm_fields=number,resolved_at,resolved_by,knowledge,close_code,close_notes"
        inc_data = get_response(INC_TABLE_SPEC, tracker.sender_id, query_filter)

        if inc_data:
            resolved_at = inc_data['resolved_at']
            resolved_by = inc_data['resolved_by']
            knowledge = inc_data['knowledge']
            resolution_code = inc_data['close_code']
            resolution_notes = inc_data['close_notes']

            query_filter = "^sla.targetINresolution^sla.typeINSLA&sysparm_display_value=True"\
                "&sysparm_exclude_reference_link=True&sysparm_fields=has_breached,business_duration"
            sla_data = get_response(SLA_TABLE_SPEC, tracker.sender_id, query_filter)
            if sla_data is not None:
                has_breached = sla_data[0]['has_breached']
                business_duration = sla_data[0]['business_duration']
                if has_breached:
                    sla_met = 'Yes'
                else:
                    sla_met = 'No'

                if resolved_at == "":
                    reply = "*This incident is not resolved yet.*"
                else:
                    if knowledge=="true":
                        knowledge_article = "<br><br> *This resolution has been submitted to be published as knowledge article.*"
                    else:
                        knowledge_article = ""

                    reply = "Following is the summary of resolution of the incident: - "\
                        "<br> **Resolved At:** "+resolved_at+"<br> **Resolved By:** "+resolved_by+\
                        "<br> **SLA met:** "+sla_met+"<br> **Resolution Duration:** "+business_duration+\
                        "<br> **Resolution Code:** "+resolution_code+"<br> **Resolution Notes:** "\
                        +resolution_notes+knowledge_article

                dispatcher.utter_message(text = reply)

        return []

class ActionAttachment(Action):
    """Fetches Files Attached to the Incident from Attachment Table in ITSM"""

    def name(self) -> Text:
        return "action_attachment"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        data = get_attachment(tracker.sender_id)
        if data is not None:
            if len(data) == 0:
                dispatcher.utter_message(text = "There are no attachments for this incident.")
            else:
                attachment_links = ''
                for attachment in data:
                    attachment_links = attachment_links+' - ['+attachment['file_name']+']('+attachment['download_link']+')\n'

                reply = "Following are the files attached to the incident:\n"\
                    +attachment_links+"\n*(Click on the ones you wish to download.)*"

                dispatcher.utter_message(text = reply)

        return []

################# Reminder Actions #################
class ActionReactToSlaReminder(Action):
    """Sends reminder for SLA breach on receiving trigger from ITSM"""

    def name(self) -> Text:
        return "action_react_to_sla_reminder"

    async def run(
        self,
        dispatcher: CollectingDispatcher,
        tracker: Tracker,
        domain: Dict[Text, Any],
    ) -> List[Dict[Text, Any]]:

        time_left = tracker.get_slot('time_left')
        percentage = tracker.get_slot('percentage')
        chat_id = tracker.sender_id

        bot_conv_url =  BOT_URL+chat_id+"/activities"
        bot_headers = get_bot_headers()

        if int(time_left.split(' ')[0]) == 0:
            alert = "*The sla has been breached!!!*"
        else:
            alert = "**"+time_left+"** left for Incident Resolution. **"+percentage+"%** of SLA has been elapsed!"

        sla_alert = {
            "type": "message",
            "attachments": [
                {
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "content":
                    {
                        "type": "AdaptiveCard",
                        "body": [
                            {
                                "type": "ColumnSet",
                                "columns": [
                                    {
                                        "type": "Column",
                                        "items": [
                                            {
                                                "type": "Image",
                                                "url": "https://icons.iconarchive.com/icons/paomedia/small-n-flat/48/bell-icon.png",
                                                "height": "30px",
                                                "style": "Person"
                                            }
                                        ],
                                        "width": "auto"
                                    },
                                    {
                                        "type": "Column",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": alert,
                                                "color": "Warning",
                                                "wrap": True
                                            }
                                        ],
                                        "width": "stretch"
                                    }
                                ]
                            }
                        ],
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "version": "1.2"
                    }
                }
            ]
        }

        sla_alert_response = requests.post(bot_conv_url, headers=bot_headers, data=json.dumps(sla_alert))

        if not sla_alert_response.ok:
            print("Error trying to send botframework message. Response: %s",sla_alert_response.text)

        return []

class ActionReactToNoUpdateReminder(Action):
    """Sends Reminder for updating the ticket in fixed intervals"""

    def name(self) -> Text:
        return "action_react_to_no_update_reminder"

    async def run(
        self,
        dispatcher: CollectingDispatcher,
        tracker: Tracker,
        domain: Dict[Text, Any],
    ) -> List[Dict[Text, Any]]:

        chat_id = tracker.get_slot("chat_id")
        number = tracker.get_slot("number")
        interval = int(tracker.get_slot("no_upd_interval"))

        bot_conv_url =  BOT_URL+chat_id+"/activities"
        bot_headers = get_bot_headers()

        no_update_alert = {
            "type":"message",
            "attachments": [
                {
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "content":
                    {
                        "type": "AdaptiveCard",
                        "body": [
                            {
                                "type": "ColumnSet",
                                "columns": [
                                    {
                                        "type": "Column",
                                        "items": [
                                            {
                                                "type": "Image",
                                                "url": "https://icons.iconarchive.com/icons/hopstarter/soft-scraps/48/Button-Reminder-icon.png",
                                                "height": "30px",
                                                "style": "Person"
                                            }
                                        ],
                                        "width": "auto"
                                    },
                                    {
                                        "type": "Column",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "There has been No Update on the incident since last "+str(interval)+" minutes!",
                                                "color": "Warning",
                                                "wrap": True,
                                                "weight": "Bolder"
                                            }
                                        ],
                                        "width": "stretch"
                                    }
                                ]
                            }
                        ],
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "version": "1.2"
                    }
                }
            ]
        }

        no_update_alert_response = requests.post(bot_conv_url, headers=bot_headers, data=json.dumps(no_update_alert))

        if not no_update_alert_response.ok:
            print("Error trying to send No Update Alert message. Response: %s",no_update_alert_response.text)

        ####### Resetting Reminder for No Update #######
        date = datetime.datetime.now() + datetime.timedelta(minutes=5)
        name = "no_update_"+number
        no_update_reminder = ReminderScheduled(
            "EXTERNAL_no_update_reminder",
            trigger_date_time=date,
            entities={"chat_id": chat_id, "number": number, "no_upd_interval": interval+5},
            name=name,
            kill_on_user_message=False,
        )
        print("\n\nNo Update DateTime :: ",date)
        print("~~~~~~~~~~~~No Update Reminder is Set~~~~~~~~~~~~~~")
        return [no_update_reminder]

class ActionReactToEmailReminder(Action):
    """Sends Reminder for sending Email communication in fixed intervals"""

    def name(self) -> Text:
        return "action_react_to_email_reminder"

    async def run(
        self,
        dispatcher: CollectingDispatcher,
        tracker: Tracker,
        domain: Dict[Text, Any],
    ) -> List[Dict[Text, Any]]:

        chat_id = tracker.get_slot("chat_id")

        email_alert = {
            "type": "message",
            "attachments": [
                {
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "content":
                    {
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "type": "AdaptiveCard",
                        "version": "1.2",
                        "body": [
                            {
                                "type": "ColumnSet",
                                "columns": [
                                    {
                                        "type": "Column",
                                        "width": "auto",
                                        "items": [
                                            {
                                                "type": "Image",
                                                "url": "https://icons.iconarchive.com/icons/social-media-icons/glossy-social/48/Email-icon.png"
                                            }
                                        ]
                                    },
                                    {
                                        "type": "Column",
                                        "width": "stretch",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "Reminder for Sending Email Communication to Stakeholders!",
                                                "wrap": True,
                                                "weight": "Bolder",
                                                "size": "Medium",
                                                "color": "Accent"
                                            }
                                        ]
                                    }
                                ]
                            }
                        ]
                    }
                }
            ]
        }

        bot_conv_url =  BOT_URL+chat_id+"/activities"
        bot_headers = get_bot_headers()

        email_card_response = requests.post(bot_conv_url, headers=bot_headers, data=json.dumps(email_alert))

        if not email_card_response.ok:
            print("Error trying to send email reminder alert card. Response: %s",email_card_response.text)

        ###### Set Email Reminder ######
        email_update_time = datetime.datetime.now() + datetime.timedelta(minutes=5)
        email_update_reminder = ReminderScheduled(
            "EXTERNAL_email_reminder",
            trigger_date_time=email_update_time,
            entities={"chat_id": chat_id},
            name="email_update_rem",
            kill_on_user_message=False,
        )

        print("\n\nEmail Update DateTime :: ",email_update_time)
        print("~~~~~~~~~~~~Email Update Reminder is Set~~~~~~~~~~~~~~")

        return [email_update_reminder]

############### Data Push Actions #################

class ActionUtterWorknoteCard(Action):
    """Sends card for updating the worknote"""

    def name(self) -> Text:
        return "action_utter_worknote_card"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        worknote_card = {
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
                                "text": "Worknote",
                                "size": "Medium",
                                "weight": "Bolder"
                            },
                            {
                                "type": "Input.Text",
                                "placeholder": "Enter the note.",
                                "id": "note",
                                "isMultiline": True
                            },
                            {
                                "type": "Input.Toggle",
                                "title": "Customer Visible (comments)",
                                "id": "customer_visible"
                            },
                            {
                                "type": "ActionSet",
                                "horizontalAlignment": "Center",
                                "actions": [
                                    {
                                        "type": "Action.Submit",
                                        "title": "Update",
                                        "data": {
                                            "msteams": {
                                                "type": "messageBack",
                                                "text": "User interaction with worknote card"
                                            }
                                        }
                                    }
                                ],
                                "spacing": "Small"
                            }
                        ]
                    }
                }
            ]
        }

        dispatcher.utter_message(json_message=worknote_card)

        return []

class ActionUpdateWorknote(Action):
    """Updates the worknote on receiving response from worknote card"""

    def name(self) -> Text:
        return "action_update_worknote"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        with open("card_data.txt", "r+", encoding='utf8') as file:
            value = file.readline().split('\n')[0]
            value = json.loads(value)
            user = file.readline().split('\n')[0]
            activity_id = file.readline()
            file.seek(0)
            file.truncate()

        if 'note' in value:
            wk_note = value['note']+"\n\n- Update given by: "+user+" via IIM"
            if value['customer_visible'] == 'true':
                data = {"comments": wk_note}
            else:
                data = {"work_notes": wk_note}

            query_filter = '?sysparm_fields=work_notes,comments'
            data = json.dumps(data)

            response_status = put_response(INC_TABLE_SPEC, tracker.sender_id, query_filter, data)

            if response_status:
                update_activity(tracker.sender_id,activity_id,user)

        else:
            dispatcher.utter_message(text='Please enter the note before submitting.')

        return []

class ActionUtterAssignmentCard(Action):
    """Sends card for updating the Engineer and Assignment Group"""

    def name(self) -> Text:
        return "action_utter_assignment_card"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        choice_list = get_groups()

        assignment_card = {
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
                                "text": "Ticket Assignment",
                                "size": "Medium",
                                "weight": "Bolder"
                            },
                            {
                                "type": "Input.ChoiceSet",
                                "choices": choice_list,
                                "placeholder": "Select a group",
                                "id": "group"
                            },
                            {
                                "type": "Input.Text",
                                "placeholder": "Enter engineer name or user id",
                                "id": "engineer"
                            },
                            {
                                "type": "ActionSet",
                                "horizontalAlignment": "Center",
                                "actions": [
                                    {
                                        "type": "Action.Submit",
                                        "title": "Update",
                                        "data": {
                                            "msteams": {
                                                "type": "messageBack",
                                                "text": "User interaction with assignment card"
                                            }
                                        }
                                    }
                                ],
                                "spacing": "Small"
                            }
                        ]
                    }
                }
            ]
        }

        dispatcher.utter_message(json_message=assignment_card)

        return []

class ActionUpdateAssignment(Action):
    """Updates the assignment fields on receiving response from assignment card"""

    def name(self) -> Text:
        return "action_update_assignment"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        with open("card_data.txt", "r+", encoding='utf8') as file:
            value = file.readline().split('\n')[0]
            value = json.loads(value)
            user = file.readline().split('\n')[0]
            activity_id = file.readline()
            file.seek(0)
            file.truncate()

        if 'engineer' in value:
            data = {
                    "assigned_to": value['engineer']
                }
            if 'group' in value:
                work_note = "The assignment group has been changed to "+ value['group']\
                    +" and the ticket has been assigned to "+value['engineer'] \
                    +"\n\n- Update given by: "+user+" via IIM"
                data["assignment_group"] = value['group']
            else:
                work_note = "The ticket has been assigned to "+ value['engineer'] + "\n\n- Update given by: "+user+" via IIM"
            data["work_notes"] = work_note

            query_filter = "?sysparm_display_value=True&sysparm_exclude_reference_link=True"\
                "&sysparm_fields=assigned_to,assignment_group"

        elif 'group' in value:
            work_note = "The assignment group has been changed to "+ value['group']+"\n\n- Update given by: "+user+" via IIM"
            data = {
                "assignment_group": value['group'],
                "work_notes": work_note
            }
            query_filter = '?sysparm_display_value=True&sysparm_exclude_reference_link=True&sysparm_fields=assignment_group'

        else:
            dispatcher.utter_message(text='Please mention the group and/or engineer you want to assign the ticket to.')
            return []

        response_status = put_response(INC_TABLE_SPEC, tracker.sender_id, query_filter, json.dumps(data))

        if response_status == 403:
            dispatcher.utter_message(text="The entered Engineer either does not exist or does not belong to the current Assignment Group.")
            dispatcher.utter_message(text="*Change the assignment group first if the engineer belongs to another group.")
            dispatcher.utter_message(text="Ticket update cancelled.")
        elif response_status:
            update_activity(tracker.sender_id,activity_id,user)

        return []

class ActionUtterStateCard(Action):
    """Sends card for updating the Incident State"""

    def name(self) -> Text:
        return "action_utter_state_card"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        query_filter = '?sysparm_display_value=True&sysparm_fields=state'
        data = get_response(INC_TABLE_SPEC, tracker.sender_id, query_filter)
        current_state = data['state']

        state_card = {
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
                                "text": "Update Ticket State",
                                "size": "Medium",
                                "weight": "Bolder"
                            },
                            {
                                "type": "TextBlock",
                                "text": "Select a state: ",
                                "spacing": "ExtraLarge"
                            },
                            {
                                "type": "ActionSet",
                                "actions": [
                                    {
                                        "type": "Action.Submit",
                                        "title": "In Progress",
                                        "data": {
                                            "msteams": {
                                                "type": "messageBack",
                                                "text": "User interaction with state card : In Progress"
                                            }
                                        }
                                    },
                                    {
                                        "type": "Action.ShowCard",
                                        "title": "On Hold",
                                        "card": {
                                            "type": "AdaptiveCard",
                                            "body": [
                                                {
                                                    "type": "ColumnSet",
                                                    "columns": [
                                                        {
                                                            "type": "Column",
                                                            "width": "auto",
                                                            "items": [
                                                                {
                                                                    "type": "TextBlock",
                                                                    "text": "On Hold Reason"
                                                                }
                                                            ]
                                                        },
                                                        {
                                                            "type": "Column",
                                                            "width": "auto",
                                                            "items": [
                                                                {
                                                                    "type": "TextBlock",
                                                                    "text": r"\*",
                                                                    "color": "Warning"
                                                                }
                                                            ],
                                                            "spacing": "None"
                                                        }
                                                    ]
                                                },
                                                {
                                                    "type": "Input.ChoiceSet",
                                                    "choices": [
                                                        {
                                                            "title": "Awaiting Caller",
                                                            "value": "Awaiting Caller"
                                                        },
                                                        {
                                                            "title": "Awaiting Change",
                                                            "value": "Awaiting Change"
                                                        },
                                                        {
                                                            "title": "Awaiting Evidence",
                                                            "value": "Awaiting Evidence"
                                                        },
                                                        {
                                                            "title": "Awaiting Problem",
                                                            "value": "Awaiting Problem"
                                                        },
                                                        {
                                                            "title": "Awaiting Vendor",
                                                            "value": "Awaiting Vendor"
                                                        }
                                                    ],
                                                    "id": "reason",
                                                    "placeholder": "Select a reason"
                                                },
                                                {
                                                    "type": "TextBlock",
                                                    "text": "Additional Comment (Only if Awaiting Caller)"
                                                },
                                                {
                                                    "type": "Input.Text",
                                                    "placeholder": "Enter note for caller",
                                                    "id": "comment",
                                                    "isMultiline": True
                                                },
                                                {
                                                    "type": "ActionSet",
                                                    "horizontalAlignment": "Center",
                                                    "actions": [
                                                        {
                                                            "type": "Action.Submit",
                                                            "title": "Confirm On Hold State",
                                                            "data": {
                                                                "msteams": {
                                                                    "type": "messageBack",
                                                                    "text": "User interaction with state card : On Hold"
                                                                }
                                                            }
                                                        }
                                                    ],
                                                    "spacing": "Small"
                                                }
                                            ]
                                        }
                                    },
                                    {
                                        "type": "Action.ShowCard",
                                        "title": "Resolved",
                                        "card": {
                                            "type": "AdaptiveCard",
                                            "body": [
                                                {
                                                    "type": "ColumnSet",
                                                    "columns": [
                                                        {
                                                            "type": "Column",
                                                            "width": "auto",
                                                            "items": [
                                                                {
                                                                    "type": "TextBlock",
                                                                    "text": "Resolution Code"
                                                                }
                                                            ]
                                                        },
                                                        {
                                                            "type": "Column",
                                                            "width": "auto",
                                                            "items": [
                                                                {
                                                                    "type": "TextBlock",
                                                                    "text": r"\*",
                                                                    "color": "Warning"
                                                                }
                                                            ],
                                                            "spacing": "None"
                                                        }
                                                    ]
                                                },
                                                {
                                                    "type": "Input.ChoiceSet",
                                                    "choices": [
                                                        {
                                                            "title": "Solved (Work Around)",
                                                            "value": "Solved (Work Around)"
                                                        },
                                                        {
                                                            "title": "Solved (Permanently)",
                                                            "value": "Solved (Permanently)"
                                                        },
                                                        {
                                                            "title": "Solved Remotely (Work Around)",
                                                            "value": "Solved Remotely (Work Around)"
                                                        },
                                                        {
                                                            "title": "Solved Remotely (Permanently)",
                                                            "value": "Solved Remotely (Permanently)"
                                                        },
                                                        {
                                                            "title": "Not Solved (Not Reproducible)",
                                                            "value": "Not Solved (Not Reproducible)"
                                                        },
                                                        {
                                                            "title": "Not Solved (Too Costly)",
                                                            "value": "Not Solved (Too Costly)"
                                                        },
                                                        {
                                                            "title": "Closed/Resolved by Caller",
                                                            "value": "Closed/Resolved by Caller"
                                                        }
                                                    ],
                                                    "id": "code",
                                                    "placeholder": "Select a code"
                                                },
                                                {
                                                    "type": "ColumnSet",
                                                    "columns": [
                                                        {
                                                            "type": "Column",
                                                            "width": "auto",
                                                            "items": [
                                                                {
                                                                    "type": "TextBlock",
                                                                    "text": "Close Notes"
                                                                }
                                                            ]
                                                        },
                                                        {
                                                            "type": "Column",
                                                            "width": "auto",
                                                            "items": [
                                                                {
                                                                    "type": "TextBlock",
                                                                    "text": r"\*",
                                                                    "color": "Warning"
                                                                }
                                                            ],
                                                            "spacing": "None"
                                                        }
                                                    ]
                                                },
                                                {
                                                    "type": "Input.Text",
                                                    "placeholder": "Enter Close notes",
                                                    "id": "close_notes",
                                                    "isMultiline": True
                                                },
                                                {
                                                    "type": "ActionSet",
                                                    "horizontalAlignment": "Center",
                                                    "actions": [
                                                        {
                                                            "type": "Action.Submit",
                                                            "title": "Confirm Resolve State",
                                                            "data": {
                                                            "msteams": {
                                                                "type": "messageBack",
                                                                "text": "User interaction with state card : Resolved"
                                                            }
                                                        }
                                                        }
                                                    ],
                                                    "spacing": "Small"
                                                }
                                            ]
                                        }
                                    }
                                ],
                                "spacing": "None"
                            },
                            {
                                "type": "TextBlock",
                                "text": "**Current state is '"+current_state+"'",
                                "fontType": "Monospace"
                            }
                        ]
                    }
                }
            ]
        }

        dispatcher.utter_message(json_message=state_card)
        return []

class ActionUpdateState(Action):
    """Updates the incident state on receiving response from state card"""

    def name(self) -> Text:
        return "action_update_state"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        with open("card_data.txt", "r+", encoding='utf8') as file:
            value = file.readline().split('\n')[0]
            value = json.loads(value)
            user = file.readline().split('\n')[0]
            activity_id = file.readline()
            file.seek(0)
            file.truncate()

        query_filter = '?sysparm_display_value=True&sysparm_fields=state'
        data = get_response(INC_TABLE_SPEC, tracker.sender_id, query_filter)
        from_state = data['state']
        to_state = tracker.get_slot('new_state')

        if to_state == from_state:
            dispatcher.utter_message(text="The incident is already in the selected state.")
            return []

        work_note = "The ticket state has been changed from *" + from_state + "* to *"\
            + to_state + "* \n\n- Update given by: "+user+" via IIM"
        data = {
            "state": to_state,
            "work_notes": work_note
        }
        if to_state == 'On Hold':
            if 'reason' in value:
                reason = value['reason']
                data['hold_reason'] = reason
                if reason == 'Awaiting Caller':
                    if 'comment' in value:
                        data['comments'] = value['comment']
                    else:
                        dispatcher.utter_message(
                            text="Please enter the note for Caller if putting the ticket on hold and "\
                                "reason is Awaiting Caller"
                        )
                        return []
            else:
                dispatcher.utter_message(text="Please select the reason for changing the ticket state to 'On Hold'")
                return []
        elif to_state == "Resolved":
            if value.keys() >= {"code","close_notes"}:
                data['close_code'] = value['code']
                data['close_notes'] = value['close_notes']
            else:
                dispatcher.utter_message(text="Please fill both the mandatory fields for changing the ticket state to resolved.")
                return []

        query_filter = '?sysparm_display_value=True&sysparm_exclude_reference_link=True&sysparm_fields=state'

        response_status = put_response(INC_TABLE_SPEC, tracker.sender_id, query_filter, json.dumps(data))

        if response_status:
            update_activity(tracker.sender_id,activity_id,user)

        return []

class ActionUploadFile(Action):
    """Attaches the uploaded file to the Incident in ITSM"""

    def name(self) -> Text:
        return "action_upload_file"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        with open("card_data.txt", "r+", encoding='utf8') as file:
            content_url = file.readline().split('\n')[0]
            user = file.readline()
            file.seek(0)
            file.truncate()

        if not content_url:
            dispatcher.utter_message(text="No File Attached. Upload Failed!")
        else:
            now = datetime.datetime.now().strftime("%d%m%Y%H%M%S")
            file_data, file_extension = get_file_content(content_url)
            print("File Type: ",file_extension)
            if file_extension is None:
                dispatcher.utter_message(text="Only JPEG and PNG file types are supported.")
                return []

            file_name = "Attachment_"+now+"."+file_extension
            response = post_attachment(tracker.sender_id,file_name,file_data)
            if response is None:
                dispatcher.utter_message(text='File Upload Failed! Please try again.')
                return []

            ##### Updating the worknote #####
            worknote = "A File has been attached to the incident."+" \n\n - Uploaded by: "+user+" via IIM"
            data = {"work_notes": worknote}
            query_filter = '?sysparm_fields=work_notes'
            response_status = put_response(INC_TABLE_SPEC, tracker.sender_id, query_filter, json.dumps(data))

            if response_status not in [None, 403]:
                dispatcher.utter_message(text="The file has been attached to the incident.")

        return []

###### Email Notification Action #########

class ActionUtterEmailUpdateCard(Action):
    """Sends card for Email communication"""

    def name(self) -> Text:
        return "action_utter_email_update_card"

    async def run(
        self,
        dispatcher: CollectingDispatcher,
        tracker: Tracker,
        domain: Dict[Text, Any],
    ) -> List[Dict[Text, Any]]:

        chat_id = tracker.sender_id

        old_email_upd_card_id = tracker.get_slot("emailUpdCardId")
        old_email_det_card_id = tracker.get_slot("emailDetCardId")
        bot_headers = get_bot_headers()

        if old_email_upd_card_id is not None:
            bot_delete_url =  BOT_URL+chat_id+"/activities/"+old_email_upd_card_id
            card_delete_response = requests.delete(bot_delete_url, headers=bot_headers)
            if not card_delete_response.ok:
                print("Error trying to delete Email Update Card. Response: %s",card_delete_response.text)

        if old_email_det_card_id is not None:
            bot_delete_url =  BOT_URL+chat_id+"/activities/"+old_email_det_card_id
            card_delete_response = requests.delete(bot_delete_url, headers=bot_headers)
            if not card_delete_response.ok:
                print("Error trying to delete Email Details Card. Response: %s",card_delete_response.text)

        es_dict = {}

        for es_val in EMAIL_SLOTS:
            es_dict[es_val] = tracker.get_slot(es_val)

        subject = "MIM - "+es_dict['EImpClient']+" - "+es_dict['Epriority']+"| "+es_dict['Emistate']+" | <"\
            +es_dict['EImpApp']+"> | "+es_dict['EIncRef']+" | "+es_dict['EIncSummary']

        email_update_card = {
            "type": "message",
            "attachments": [
                {
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "content":
                    {
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "type": "AdaptiveCard",
                        "version": "1.2",
                        "body": [
                            {
                                "type": "TextBlock",
                                "text": "Give the Latest Update for Email Communication",
                                "size": "Medium",
                                "weight": "Bolder",
                                "wrap": True
                            },
                            {
                                "type": "TextBlock",
                                "text": "Email Subject:",
                                "weight": "Bolder",
                                "color": "Accent"
                            },
                            {
                                "type": "Input.Text",
                                "id": "ESub",
                                "spacing": "None",
                                "value": subject,
                                "isMultiline": True
                            },
                            {
                                "type": "Input.Text",
                                "placeholder": "Latest Update",
                                "id": "latest_update",
                                "isMultiline": True
                            },
                            {
                                "type": "ActionSet",
                                "horizontalAlignment": "Center",
                                "actions": [
                                    {
                                        "type": "Action.ShowCard",
                                        "title": "Review Details",
                                        "card": {
                                            "type": "AdaptiveCard",
                                            "body": [
                                                {
                                                    "type": "FactSet",
                                                    "facts": [
                                                        {
                                                            "title": "Priority:",
                                                            "value": es_dict['Epriority']
                                                        },
                                                        {
                                                            "title": "Major Incident State:",
                                                            "value": es_dict['Emistate']
                                                        },
                                                        {
                                                            "title": "Incident Summary:",
                                                            "value": es_dict['EIncSummary']
                                                        },
                                                        {
                                                            "title": "Business Impact:",
                                                            "value": es_dict['EBizImp']
                                                        },
                                                        {
                                                            "title": "Impacted Location(s):",
                                                            "value": es_dict['EImpLoc']
                                                        },
                                                        {
                                                            "title": "Impacted Clients:",
                                                            "value": es_dict['EImpClient']
                                                        },
                                                        {
                                                            "title": "Impacted Applications:",
                                                            "value": es_dict['EImpApp']
                                                        },
                                                        {
                                                            "title": "No. of Users Impacted:",
                                                            "value": es_dict['EImpUsr']
                                                        },
                                                        {
                                                            "title": "Issue Reported By:",
                                                            "value": es_dict['EIsRepBy']
                                                        },
                                                        {
                                                            "title": "Incident Ticket Ref#",
                                                            "value": es_dict['EIncRef']
                                                        },
                                                        {
                                                            "title": "Vendor Name:",
                                                            "value": es_dict['EVendor']
                                                        },
                                                        {
                                                            "title": "Incident Start Date/Time:",
                                                            "value": es_dict['EIncStart']
                                                        },
                                                        {
                                                            "title": "Date/Time (MIM Engaged):",
                                                            "value": es_dict['EMimEng']
                                                        },
                                                        {
                                                            "title": "Major Incident Manager:",
                                                            "value": es_dict['EMIM']
                                                        },
                                                        {
                                                            "title": "Support Teams Involved:",
                                                            "value": es_dict['ESupTeams']
                                                        },
                                                        {
                                                            "title": "Workaround",
                                                            "value": es_dict['EWrkArnd']
                                                        },
                                                        {
                                                            "title": "Change Related/Ref:",
                                                            "value": es_dict['EChange']
                                                        },
                                                        {
                                                            "title": "Reason for Outage (RFO):",
                                                            "value": es_dict['ERFO']
                                                        },
                                                        {
                                                            "title": "Actual Resolution Time:",
                                                            "value": es_dict['EResTime']
                                                        },
                                                        {
                                                            "title": "Outage Duration:",
                                                            "value": es_dict['EOutDur']
                                                        },
                                                        {
                                                            "title": "Next Update Due:",
                                                            "value": es_dict['ENxtUpd']
                                                        },
                                                        {
                                                            "title": "Distribution Lists:",
                                                            "value": es_dict['EDL']
                                                        }
                                                    ]
                                                },
                                                {
                                                    "type": "ActionSet",
                                                    "actions": [
                                                        {
                                                            "type": "Action.Submit",
                                                            "title": "Edit Details",
                                                            "data": {
                                                                "msteams": {
                                                                    "type": "messageBack",
                                                                    "text": "User clicked Edit Details button"
                                                                }
                                                            }
                                                        }
                                                    ],
                                                    "horizontalAlignment": "Center",
                                                    "spacing": "None"
                                                }
                                            ]
                                        }
                                    },
                                    {
                                        "type": "Action.Submit",
                                        "title": "Send Email",
                                        "data": {
                                            "msteams": {
                                                "type": "messageBack",
                                                "text": "User interaction with Email Update card"
                                            }
                                        }
                                    }
                                ],
                                "spacing": "Small"
                            }
                        ]
                    }
                }
            ]
        }

        bot_conv_url =  BOT_URL+chat_id+"/activities"
        bot_headers = get_bot_headers()

        email_card_response = requests.post(bot_conv_url, headers=bot_headers, data=json.dumps(email_update_card))

        if not email_card_response.ok:
            print("Error trying to send email update card. Response: %s",email_card_response.text)

        return [SlotSet("emailUpdCardId",email_card_response.json()['id']), SlotSet('emailDetCardId', None)]

class ActionUtterEmailDetailsCard(Action):
    """Sends card for updating Email Details"""

    def name(self) -> Text:
        return "action_utter_email_details_card"

    async def run(
        self,
        dispatcher: CollectingDispatcher,
        tracker: Tracker,
        domain: Dict[Text, Any],
    ) -> List[Dict[Text, Any]]:

        chat_id = tracker.sender_id

        old_email_upd_card_id = tracker.get_slot("emailUpdCardId")
        old_email_det_card_id = tracker.get_slot("emailDetCardId")
        bot_headers = get_bot_headers()

        if old_email_upd_card_id is not None:
            bot_delete_url =  BOT_URL+chat_id+"/activities/"+old_email_upd_card_id
            card_delete_response = requests.delete(bot_delete_url, headers=bot_headers)
            if not card_delete_response.ok:
                print("Error trying to delete Email Update Card. Response: %s",card_delete_response.text)

        if old_email_det_card_id is not None:
            bot_delete_url =  BOT_URL+chat_id+"/activities/"+old_email_det_card_id
            card_delete_response = requests.delete(bot_delete_url, headers=bot_headers)
            if not card_delete_response.ok:
                print("Error trying to delete Email Details Card. Response: %s",card_delete_response.text)

        es_dict = {}

        for es_val in EMAIL_SLOTS:
            es_dict[es_val] = tracker.get_slot(es_val)

        email_details_card = {
            "type": "message",
            "attachments": [
                {
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "content":
                    {
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "type": "AdaptiveCard",
                        "version": "1.2",
                        "body": [
                            {
                                "type": "TextBlock",
                                "size": "medium",
                                "weight": "bolder",
                                "text": "Email Communication Details",
                                "horizontalAlignment": "center"
                            },
                            {
                                "type": "Container",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Priority:",
                                        "color": "Accent",
                                        "weight": "Bolder"
                                    }
                                ]
                            },
                            {
                                "type": "Input.ChoiceSet",
                                "choices": [
                                    {
                                        "title": "P1",
                                        "value": "P1"
                                    },
                                    {
                                        "title": "P2",
                                        "value": "P2"
                                    }
                                ],
                                "id": "Epriority",
                                "value": es_dict['Epriority'],
                                "spacing": "None"
                            },
                            {
                                "type": "Container",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Major Incident State:",
                                        "color": "Accent",
                                        "weight": "Bolder"
                                    }
                                ]
                            },
                            {
                                "type": "Input.ChoiceSet",
                                "choices": [
                                    {
                                        "title": "Declared",
                                        "value": "Declared"
                                    },
                                    {
                                        "title": "In-Progress",
                                        "value": "In-Progress"
                                    },
                                    {
                                        "title": "On-Hold",
                                        "value": "On-Hold"
                                    },
                                    {
                                        "title": "Under Observation",
                                        "value": "Under Observation"
                                    },
                                    {
                                        "title": "Restored/Monitoring",
                                        "value": "Restored/Monitoring"
                                    },
                                    {
                                        "title": "Resolved",
                                        "value": "Resolved"
                                    },
                                    {
                                        "title": "Downgraded",
                                        "value": "Downgraded"
                                    },
                                    {
                                        "title": "Cancelled",
                                        "value": "Cancelled"
                                    }
                                ],
                                "id": "Emistate",
                                "value": es_dict['Emistate'],
                                "spacing": "None"
                            },
                            {
                                "type": "Container",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Incident Summary:",
                                        "color": "Accent",
                                        "weight": "Bolder"
                                    }
                                ]
                            },
                            {
                                "type": "Input.Text",
                                "id": "EIncSummary",
                                "spacing": "None",
                                "value": es_dict['EIncSummary']
                            },
                            {
                                "type": "Container",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Business Impact (Description):",
                                        "color": "Accent",
                                        "weight": "Bolder"
                                    }
                                ]
                            },
                            {
                                "type": "Input.Text",
                                "id": "EBizImp",
                                "spacing": "None",
                                "value": es_dict['EBizImp']
                            },
                            {
                                "type": "Container",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Impacted Location(s) / Sites:",
                                        "color": "Accent",
                                        "weight": "Bolder"
                                    }
                                ]
                            },
                            {
                                "type": "Input.Text",
                                "id": "EImpLoc",
                                "spacing": "None",
                                "value": es_dict['EImpLoc']
                            },
                            {
                                "type": "Container",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Impacted Clients:",
                                        "color": "Accent",
                                        "weight": "Bolder"
                                    }
                                ]
                            },
                            {
                                "type": "Input.Text",
                                "id": "EImpClient",
                                "spacing": "None",
                                "value": es_dict['EImpClient']
                            },
                            {
                                "type": "Container",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Impacted Application / Services:",
                                        "color": "Accent",
                                        "weight": "Bolder"
                                    }
                                ]
                            },
                            {
                                "type": "Input.Text",
                                "id": "EImpApp",
                                "spacing": "None",
                                "value": es_dict['EImpApp']
                            },
                            {
                                "type": "Container",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "No. of users impacted:",
                                        "color": "Accent",
                                        "weight": "Bolder"
                                    }
                                ]
                            },
                            {
                                "type": "Input.Text",
                                "id": "EImpUsr",
                                "spacing": "None",
                                "value": es_dict['EImpUsr']
                            },
                            {
                                "type": "Container",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Issue reported by:",
                                        "color": "Accent",
                                        "weight": "Bolder"
                                    }
                                ],
                                "separator": True
                            },
                            {
                                "type": "Input.Text",
                                "id": "EIsRepBy",
                                "spacing": "None",
                                "value": es_dict['EIsRepBy']
                            },
                            {
                                "type": "Container",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Incident Ticket Ref#",
                                        "color": "Accent",
                                        "weight": "Bolder"
                                    }
                                ]
                            },
                            {
                                "type": "Input.Text",
                                "id": "EIncRef",
                                "spacing": "None",
                                "value": es_dict['EIncRef']
                            },
                            {
                                "type": "Container",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Vendor Name / Ticket Ref:",
                                        "color": "Accent",
                                        "weight": "Bolder"
                                    }
                                ]
                            },
                            {
                                "type": "Input.Text",
                                "id": "EVendor",
                                "spacing": "None",
                                "value": es_dict['EVendor']
                            },
                            {
                                "type": "Container",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Incident Start Date/ Time:",
                                        "color": "Accent",
                                        "weight": "Bolder"
                                    }
                                ]
                            },
                            {
                                "type": "Input.Text",
                                "id": "EIncStart",
                                "spacing": "None",
                                "value": es_dict['EIncStart']
                            },
                            {
                                "type": "Container",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Date / Time (MIM Engaged):",
                                        "color": "Accent",
                                        "weight": "Bolder"
                                    }
                                ]
                            },
                            {
                                "type": "Input.Text",
                                "id": "EMimEng",
                                "spacing": "None",
                                "value": es_dict['EMimEng']
                            },
                            {
                                "type": "Container",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Major Incident Manager:",
                                        "color": "Accent",
                                        "weight": "Bolder"
                                    }
                                ],
                                "separator": True
                            },
                            {
                                "type": "Input.ChoiceSet",
                                "choices": [
                                    {
                                        "title": "MIM Mgr 1",
                                        "value": "MIM Mgr 1"
                                    },
                                    {
                                        "title": "MIM Mgr 2",
                                        "value": "MIM Mgr 2"
                                    },
                                    {
                                        "title": "MIM Mgr 3",
                                        "value": "MIM Mgr 3"
                                    }
                                ],
                                "id": "EMIM",
                                "value": es_dict['EMIM'],
                                "spacing": "None"
                            },
                            {
                                "type": "Container",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Support Teams involved:",
                                        "color": "Accent",
                                        "weight": "Bolder"
                                    }
                                ]
                            },
                            {
                                "type": "Input.Text",
                                "id": "ESupTeams",
                                "spacing": "None",
                                "value": es_dict['ESupTeams']
                            },
                            {
                                "type": "Container",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Workaround:",
                                        "color": "Accent",
                                        "weight": "Bolder"
                                    }
                                ],
                                "separator": True
                            },
                            {
                                "type": "Input.Text",
                                "id": "EWrkArnd",
                                "spacing": "None",
                                "value": es_dict['EWrkArnd']
                            },
                            {
                                "type": "Container",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Change Related / Ref:",
                                        "color": "Accent",
                                        "weight": "Bolder"
                                    }
                                ],
                                "separator": True
                            },
                            {
                                "type": "Input.Text",
                                "id": "EChange",
                                "spacing": "None",
                                "value": es_dict['EChange']
                            },
                            {
                                "type": "Container",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Reason for Outage (RFO):",
                                        "color": "Accent",
                                        "weight": "Bolder"
                                    }
                                ]
                            },
                            {
                                "type": "Input.Text",
                                "id": "ERFO",
                                "spacing": "None",
                                "value": es_dict['ERFO']
                            },
                            {
                                "type": "Container",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Actual Resolution Time:",
                                        "color": "Accent",
                                        "weight": "Bolder"
                                    }
                                ]
                            },
                            {
                                "type": "Input.Text",
                                "id": "EResTime",
                                "spacing": "None",
                                "value": es_dict['EResTime']
                            },
                            {
                                "type": "Container",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Outage Duration:",
                                        "color": "Accent",
                                        "weight": "Bolder"
                                    }
                                ]
                            },
                            {
                                "type": "Input.Text",
                                "id": "EOutDur",
                                "spacing": "None",
                                "value": es_dict['EOutDur']
                            },
                            {
                                "type": "Container",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Next Update Due:",
                                        "color": "Accent",
                                        "weight": "Bolder"
                                    }
                                ]
                            },
                            {
                                "type": "Input.Text",
                                "id": "ENxtUpd",
                                "spacing": "None",
                                "value": es_dict['ENxtUpd']
                            },
                            {
                                "type": "Container",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Distribution Lists: ",
                                        "color": "Accent",
                                        "weight": "Bolder"
                                    }
                                ],
                                "separator": True
                            },
                            {
                                "type": "Input.Text",
                                "id": "EDL",
                                "spacing": "None",
                                "value": es_dict['EDL'],
                                "isMultiline": True
                            },
                            {
                                "type": "ActionSet",
                                "actions": [
                                    {
                                        "type": "Action.Submit",
                                        "title": "Save Details",
                                        "data": {
                                            "msteams": {
                                                "type": "messageBack",
                                                "text": "User interaction with Email Details card"
                                            },
                                            "init_com": "false"
                                        }
                                    }
                                ],
                                "horizontalAlignment": "Center",
                                "spacing": "None"
                            }
                        ]
                    }
                }
            ]
        }

        bot_conv_url =  BOT_URL+chat_id+"/activities"
        bot_headers = get_bot_headers()

        email_card_response = requests.post(bot_conv_url, headers=bot_headers, data=json.dumps(email_details_card))

        if not email_card_response.ok:
            print("Error trying to send email details card. Response: %s",email_card_response.text)

        return [SlotSet("emailDetCardId",email_card_response.json()['id']), SlotSet('emailUpdCardId', None)]

class ActionSaveEmailDetails(Action):
    """Stores the email details card values in slots"""

    def name(self) -> Text:
        return "action_save_email_details"

    async def run(
        self, dispatcher, tracker: Tracker, domain: Dict[Text, Any]
    ) -> List[Dict[Text, Any]]:

        chat_id = tracker.sender_id

        with open("card_data.txt", "r+", encoding='utf8') as file:
            value = file.readline().split('\n')[0]
            value = json.loads(value)
            user = file.readline().split('\n')[0]
            activity_id = file.readline()
            file.seek(0)
            file.truncate()

        if value['Emistate'] != 'Resolved':
            try:
                next_update = datetime.datetime.strptime(value['ENxtUpd'],'%d-%b-%Y %H:%M')
                if next_update < datetime.datetime.now():
                    dispatcher.utter_message(text="Please enter a future time in the Next Update Field.")
                    return []
            except ValueError:
                dispatcher.utter_message(text="Please enter the Next Update in proper format (For eg. 13-Dec-2021 01:30)")
                return []

        return_events_list = []
        try:
            email_events = [
                SlotSet('Epriority', value['Epriority']),
                SlotSet('Emistate', value['Emistate']),
                SlotSet('EIncSummary', value['EIncSummary']),
                SlotSet('EBizImp', value['EBizImp']),
                SlotSet('EImpLoc', value['EImpLoc']),
                SlotSet('EImpClient', value['EImpClient']),
                SlotSet('EImpApp', value['EImpApp']),
                SlotSet('EImpUsr', value['EImpUsr']),
                SlotSet('EIsRepBy', value['EIsRepBy']),
                SlotSet('EIncRef', value['EIncRef']),
                SlotSet('EVendor', value['EVendor']),
                SlotSet('EIncStart', value['EIncStart']),
                SlotSet('EMimEng', value['EMimEng']),
                SlotSet('EMIM', value['EMIM']),
                SlotSet('ESupTeams', value['ESupTeams']),
                SlotSet('EWrkArnd', value['EWrkArnd']),
                SlotSet('EChange', value['EChange']),
                SlotSet('ERFO', value['ERFO']),
                SlotSet('EResTime', value['EResTime']),
                SlotSet('EOutDur', value['EOutDur']),
                SlotSet('ENxtUpd', value['ENxtUpd']),
                SlotSet('EDL', value['EDL'])
            ]
        except:
            dispatcher.utter_message(text="Error trying to save the details!")
            dispatcher.utter_message(text="Please make sure that all the values are filled.")
            return []

        return_events_list.extend(email_events)

        if value['init_com'] == "true":

            subject = "MIM - "+value['EImpClient']+" - "+value['Epriority']+"| "+value['Emistate']+" | <"\
                +value['EImpApp']+"> | "+value['EIncRef']+" | "+value['EIncSummary']

            email_update_card = {
                "type": "message",
                "attachments": [
                    {
                        "contentType": "application/vnd.microsoft.card.adaptive",
                        "content":
                        {
                            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                            "type": "AdaptiveCard",
                            "version": "1.2",
                            "body": [
                                {
                                    "type": "TextBlock",
                                    "text": "Give the First Update for Email Communication for Declared Major Incident",
                                    "size": "Medium",
                                    "weight": "Bolder",
                                    "wrap": True
                                },
                                {
                                    "type": "TextBlock",
                                    "text": "Email Subject:",
                                    "weight": "Bolder",
                                    "color": "Accent"
                                },
                                {
                                    "type": "Input.Text",
                                    "id": "ESub",
                                    "spacing": "None",
                                    "value": subject,
                                    "isMultiline": True
                                },
                                {
                                    "type": "Input.Text",
                                    "placeholder": "Latest Update",
                                    "id": "latest_update",
                                    "isMultiline": True
                                },
                                {
                                    "type": "ActionSet",
                                    "horizontalAlignment": "Center",
                                    "actions": [
                                        {
                                            "type": "Action.ShowCard",
                                            "title": "Review Details",
                                            "card": {
                                                "type": "AdaptiveCard",
                                                "body": [
                                                    {
                                                        "type": "FactSet",
                                                        "facts": [
                                                            {
                                                                "title": "Priority:",
                                                                "value": value['Epriority']
                                                            },
                                                            {
                                                                "title": "Major Incident State:",
                                                                "value": value['Emistate']
                                                            },
                                                            {
                                                                "title": "Incident Summary:",
                                                                "value": value['EIncSummary']
                                                            },
                                                            {
                                                                "title": "Business Impact:",
                                                                "value": value['EBizImp']
                                                            },
                                                            {
                                                                "title": "Impacted Location(s):",
                                                                "value": value['EImpLoc']
                                                            },
                                                            {
                                                                "title": "Impacted Clients:",
                                                                "value": value['EImpClient']
                                                            },
                                                            {
                                                                "title": "Impacted Applications:",
                                                                "value": value['EImpApp']
                                                            },
                                                            {
                                                                "title": "No. of Users Impacted:",
                                                                "value": value['EImpUsr']
                                                            },
                                                            {
                                                                "title": "Issue Reported By:",
                                                                "value": value['EIsRepBy']
                                                            },
                                                            {
                                                                "title": "Incident Ticket Ref#",
                                                                "value": value['EIncRef']
                                                            },
                                                            {
                                                                "title": "Vendor Name:",
                                                                "value": value['EVendor']
                                                            },
                                                            {
                                                                "title": "Incident Start Date/Time:",
                                                                "value": value['EIncStart']
                                                            },
                                                            {
                                                                "title": "Date/Time (MIM Engaged):",
                                                                "value": value['EMimEng']
                                                            },
                                                            {
                                                                "title": "Major Incident Manager:",
                                                                "value": value['EMIM']
                                                            },
                                                            {
                                                                "title": "Support Teams Involved:",
                                                                "value": value['ESupTeams']
                                                            },
                                                            {
                                                                "title": "Workaround",
                                                                "value": value['EWrkArnd']
                                                            },
                                                            {
                                                                "title": "Change Related/Ref:",
                                                                "value": value['EChange']
                                                            },
                                                            {
                                                                "title": "Reason for Outage (RFO):",
                                                                "value": value['ERFO']
                                                            },
                                                            {
                                                                "title": "Actual Resolution Time:",
                                                                "value": value['EResTime']
                                                            },
                                                            {
                                                                "title": "Outage Duration:",
                                                                "value": value['EOutDur']
                                                            },
                                                            {
                                                                "title": "Next Update Due:",
                                                                "value": value['ENxtUpd']
                                                            },
                                                            {
                                                                "title": "Distribution Lists:",
                                                                "value": value['EDL']
                                                            }
                                                        ]
                                                    },
                                                    {
                                                        "type": "ActionSet",
                                                        "actions": [
                                                            {
                                                                "type": "Action.Submit",
                                                                "title": "Edit Details",
                                                                "data": {
                                                                    "msteams": {
                                                                        "type": "messageBack",
                                                                        "text": "User clicked Edit Details button"
                                                                    }
                                                                }
                                                            }
                                                        ],
                                                        "horizontalAlignment": "Center",
                                                        "spacing": "None"
                                                    }
                                                ]
                                            }
                                        },
                                        {
                                            "type": "Action.Submit",
                                            "title": "Send Email",
                                            "data": {
                                                "msteams": {
                                                    "type": "messageBack",
                                                    "text": "User interaction with Email Update card"
                                                }
                                            }
                                        }
                                    ],
                                    "spacing": "Small"
                                }
                            ]
                        }
                    }
                ]
            }

            bot_conv_url =  BOT_URL+chat_id+"/activities"
            bot_headers = get_bot_headers()

            email_card_response = requests.post(bot_conv_url, headers=bot_headers, data=json.dumps(email_update_card))

            if not email_card_response.ok:
                print("Error trying to send email update card. Response: %s",email_card_response.text)

            return_events_list.extend([SlotSet("emailUpdCardId",email_card_response.json()['id'])])

            if value['Erem_flag'] == "true":
                return_events_list.extend([SlotSet('Erem_flag','true')])
                email_update_time = datetime.datetime.now() + datetime.timedelta(minutes=2)
                email_update_reminder = ReminderScheduled(
                    "EXTERNAL_email_reminder",
                    trigger_date_time=email_update_time,
                    entities={"chat_id": chat_id},
                    name="email_update_rem",
                    kill_on_user_message=False,
                )

                print("\n\nEmail Update DateTime :: ",email_update_time)
                print("~~~~~~~~~~~~Email Update Reminder is Set~~~~~~~~~~~~~~")

                dispatcher.utter_message(text="Email Notification reminder is set by "+user+". You will be reminded in 5 minutes.")
                return_events_list.extend([email_update_reminder])
            else:
                return_events_list.extend([SlotSet('Erem_flag','false')])
        else:
            if (
                value['Emistate'] != 'Resolved'
                and tracker.get_slot('Erem_flag') == 'true'
                and tracker.get_slot('ENxtUpd') != value['ENxtUpd']
                and tracker.get_slot('mi_state') == 'Accepted'
            ):
                ###### Set Email Reminder #######
                email_update_time = next_update - datetime.timedelta(minutes=5)
                email_update_reminder = ReminderScheduled(
                    "EXTERNAL_email_reminder",
                    trigger_date_time=email_update_time,
                    entities={"chat_id": chat_id},
                    name="email_update_rem",
                    kill_on_user_message=False,
                )

                print("\n\nEmail Update DateTime :: ",email_update_time)
                print("~~~~~~~~~~~~Email Update Reminder is Set~~~~~~~~~~~~~~")
                return_events_list.extend([email_update_reminder])

        return_events_list.extend([SlotSet('emailDetCardId',None)])
        update_activity(chat_id,activity_id,user,card_content = "Email Details have been saved.")

        return return_events_list

class ActionSendEmail(Action):
    """Sends email on receiving response from Email Update Card"""

    def name(self) -> Text:
        return "action_send_email"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        chat_id = tracker.sender_id
        return_events_list = []

        with open("card_data.txt", "r+", encoding='utf8') as file:
            value = file.readline().split('\n')[0]
            value = json.loads(value)
            user = file.readline().split('\n')[0]
            activity_id = file.readline()
            file.seek(0)
            file.truncate()

        if value.keys() >= {"latest_update", "ESub"}:
            latest_update = value['latest_update']
            subject = value['ESub']
        else:
            dispatcher.utter_message(text='Please enter values in both fields before submitting.')
            return []

        ##### Formatting the update in email template #####
        email_update_history = tracker.get_slot("emailUpdateHistory")
        current_timestamp = datetime.datetime.now().strftime("%d-%b-%Y %H:%M")
        time_stamp = "<td id=\"time_stamp\"><u>"+current_timestamp+"</u></td>"
        update = "<td id=\"update\">"+ latest_update +"</td>"
        update_row = "<tr>"+time_stamp+update+"</tr>"
        consolidated_update = update_row+email_update_history

        ##### Extracting Email Details from Slots #####
        es_dict = {}

        for es_val in EMAIL_SLOTS:
            es_dict[es_val] = tracker.get_slot(es_val)

        ######## Store the update in the Database #########
        query = f"INSERT INTO email_updates VALUES ('{tracker.get_slot('number')}',"\
            f"'{current_timestamp}','{latest_update}','{es_dict['EMIM']}')"
        ENGINE.execute(query)

        ####### Send Email Notification #######
        print("****** SEND EMAIL ******")
        mail_sent = send_email_notification(tracker.sender_id, consolidated_update, subject, es_dict)

        if mail_sent == "Success":
            update_activity(chat_id,activity_id,user,card_content = "Email Sent to Stakeholders.")
        elif mail_sent == "No DL":
            dispatcher.utter_message("Distribution Lists not entered properly.")
            return []
        else:
            dispatcher.utter_message("Some problem occurred while sending the Email.")
            return []

        nxt_upd_due = datetime.datetime.strptime(es_dict['ENxtUpd'],'%d-%b-%Y %H:%M')
        # next_update = datetime.datetime.now() + datetime.timedelta(minutes=30)
        # nxt_upd_due = next_update.strftime("%d-%b-%Y %H:%M")
        # return_events_list.extend([SlotSet('ENxtUpd', es_dict['ENxtUpd'])])
        return_events_list.extend([SlotSet("emailUpdateHistory", consolidated_update), SlotSet('emailUpdCardId', None)])

        if es_dict["Emistate"] == 'Declared' and tracker.get_slot('Erem_flag') == 'true' and tracker.get_slot('mi_state') == 'Accepted':
            ###### Set Email Reminder #######
            email_update_time = nxt_upd_due - datetime.timedelta(minutes=5)
            email_update_reminder = ReminderScheduled(
                "EXTERNAL_email_reminder",
                trigger_date_time=email_update_time,
                entities={"chat_id": chat_id},
                name="email_update_rem",
                kill_on_user_message=False,
            )

            print("\n\nEmail Update DateTime :: ",email_update_time)
            print("~~~~~~~~~~~~Email Update Reminder is Set~~~~~~~~~~~~~~")
            return_events_list.extend([email_update_reminder])

        return return_events_list

class ActionSetEmailReminder(Action):
    """Sets email reminder"""

    def name(self) -> Text:
        return "action_set_email_reminder"

    async def run(
        self, dispatcher, tracker: Tracker, domain: Dict[Text, Any]
    ) -> List[Dict[Text, Any]]:

        mi_state = tracker.get_slot('mi_state')
        if mi_state == 'Canceled':
            dispatcher.utter_message(text = 'Cannot set reminder as this is not a Major Incident')
            return []

        erem = tracker.get_slot('Erem_flag')
        if erem == 'true':
            dispatcher.utter_message(text="Email Reminder is already set.")
            return []

        ###### Set Email Reminder #######
        next_update = datetime.datetime.strptime(tracker.get_slot('ENxtUpd'),'%d-%b-%Y %H:%M')
        email_update_time = next_update - datetime.timedelta(minutes=5)
        email_update_reminder = ReminderScheduled(
            "EXTERNAL_email_reminder",
            trigger_date_time=email_update_time,
            entities={"chat_id": tracker.sender_id},
            name="email_update_rem",
            kill_on_user_message=False,
        )

        print("\n\nEmail Update DateTime :: ",email_update_time)
        print("~~~~~~~~~~~~Email Update Reminder is Set~~~~~~~~~~~~~~")

        dispatcher.utter_message(
            text="Email Reminder has been set! "\
                "You will recieve the reminder 5 mins prior to the next scheduled email update."
        )
        return [SlotSet("Erem_flag","true"), email_update_reminder]

class ActionCancelEmailReminder(Action):
    """Cancels email reminder"""

    def name(self) -> Text:
        return "action_cancel_email_reminder"

    async def run(
        self, dispatcher, tracker: Tracker, domain: Dict[Text, Any]
    ) -> List[Dict[Text, Any]]:

        dispatcher.utter_message("The Email Reminder is Cancelled!.")

        print("\n\n Cancelling Email Reminder.....")

        return [ReminderCancelled("email_update_rem")]

###### MIR Generation Action #########

class ActionGenerateMIR(Action):
    """Populates MIR Template with values and upload it to Sharepoint Site"""

    def name(self) -> Text:
        return "action_generate_MIR"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        mi_state = tracker.get_slot('mi_state')
        if mi_state == 'Canceled':
            dispatcher.utter_message(text = 'This incident is not a major incident. Hence, Report cannot be generated!')
            return []

        template = "MIR_Template.docx"
        document = MailMerge(template)

        es_dict = {}
        for es_val in EMAIL_SLOTS:
            es_dict[es_val] = tracker.get_slot(es_val)

        query_filter = '?sysparm_display_value=True&sysparm_exclude_reference_link=True&sysparm_fields=description,problem_id'
        data = get_response(INC_TABLE_SPEC, tracker.sender_id, query_filter)
        if data is not None:
            description = data['description']
            problem_ref = data['problem_id']
        else:
            description = ""
            problem_ref = "NA"

        if not problem_ref:
            problem_ref = "NA"

        document.merge(
            incident_summary=es_dict['EIncSummary'],
            incident_description=description,
            business_impact=es_dict['EBizImp'],
            impacted_sites=es_dict['EImpLoc'],
            impacted_clients=es_dict['EImpClient'],
            impacted_apps=es_dict['EImpApp'],
            no_of_users_impacted=es_dict['EImpUsr'],
            issue_reported_by=es_dict['EIsRepBy'],
            incident_reference=es_dict['EIncRef'],
            incident_priority=es_dict['Epriority'],
            vendor=es_dict['EVendor'],
            problem_ref=problem_ref,
            incident_start=es_dict['EIncStart'],
            mim_engaged=es_dict['EMimEng'],
            mi_manager=es_dict['EMIM'],
            support_teams=es_dict['ESupTeams'],
            workaround=es_dict['EWrkArnd'],
            change_related=es_dict['EChange'],
            rfo=es_dict['ERFO'],
            resolution_time=es_dict['EResTime'],
            outage_duration=es_dict['EOutDur']
        )

        ###### Get updates from Database ##########
        query = f"SELECT * from email_updates where number='{tracker.get_slot('number')}'"
        result = ENGINE.execute(query)
        sig_events = []
        for res_row in result:
            table_row = {
                'date_time': res_row['date_time'],
                'description': res_row['update'],
                'manager': res_row['manager']
            }
            sig_events.append(table_row)

        document.merge_rows('date_time', sig_events)

        document.write('MIR-output.docx')

        with open('MIR-output.docx', 'rb') as file:
            file_content = file.read()

        number = tracker.get_slot('number')
        file_name = 'MIR_' + number + '.docx'
        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drive/items/root:/{FOLDER}/{file_name}:/content"

        while True:
            response = requests.put(url, headers=actions.globals.HEADERS, data=file_content)
            if response.status_code == 401:
                print("Token Expired.....Refreshing Token")
                refresh_token()
            elif response.ok:
                web_url = response.json()['webUrl']
                download_url = response.json()['@microsoft.graph.downloadUrl'].split('?')[0]
                download_url = download_url + "?SourceUrl=https://"+TENANT+".sharepoint.com/sites/"+SITE_NAME+"/Shared%20Documents/"+FOLDER+"/"+file_name
                print("*****File Uploaded Successfully*****")
                break
            else:
                print('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())
                return []

        mir_card = {
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
                                "text": "Major Incident Report has been Shared onto the Sharepoint Site. ",
                                "wrap": True,
                                "size": "medium",
                                "weight": "bolder"
                            },
                            {
                                "type": "ActionSet",
                                "actions": [
                                    {
                                        "type": "Action.OpenUrl",
                                        "title": "View Report",
                                        "url": web_url
                                    },
                                    {
                                        "type": "Action.OpenUrl",
                                        "title": "Download Report",
                                        "url": download_url
                                    }
                                ],
                                "horizontalAlignment": "Center"
                            }
                        ]
                    }
                }
            ]
        }

        dispatcher.utter_message(json_message=mir_card)

        return []

class ActionSendMIR(Action):
    """Extracts MIR froms sharepoint site and sends to stakeholders via Email"""

    def name(self) -> Text:
        return "action_send_MIR"

    async def run(self,
           dispatcher: CollectingDispatcher,
           tracker: Tracker,
           domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        mi_state = tracker.get_slot('mi_state')
        if mi_state == 'Canceled':
            dispatcher.utter_message(text = 'This incident is not a major incident. Hence, Report cannot be sent!')
            return []

        ## Get sharepoint file ##
        file_name = 'MIR_' + tracker.get_slot('number') + '.docx'
        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drive/items/root:/General/{file_name}:/content"

        while True:
            response = requests.get(url=url, headers=actions.globals.HEADERS)
            if response.status_code == 401:
                print("Token Expired.....Refreshing Token")
                refresh_token()
            elif response.ok:
                with open('Final-MIR.docx', 'wb') as file:
                    file.write(response.content)
                break
            else:
                dispatcher.utter_message("Problem occurred while downloading the MIR from sharepoint. "\
                "Make sure that the naming is proper i.e. MIR_<INC_number>.docx")
                print('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())
                return []

        ## Send email wit MIR as attachment ##
        print("****** SEND MIR ******")
        mail_sent = send_mir(file_name)

        if mail_sent == "Success":
            dispatcher.utter_message("MIR has been sent to stakeholders via Email.")
        else:
            dispatcher.utter_message("Some problem occurred while sending the MIR.")

        return []
