"""
This file contains functions for email notification feature.
"""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.utils import formataddr

import json
import logging
from string import Template
import requests

from actions.requester import ENGINE
from actions.constants import SENDER_ADDRESS, SENDER_NAME, PASSWORD, USER, PWD, SNOW_ACCT_API_URL

logger = logging.getLogger(__name__)

# Set the request parameters
URL = SNOW_ACCT_API_URL + 'v1/email'
HEADERS = {"Content-Type":"application/json","Accept":"application/json"}

def send_email_notification(chat_id: str, updates: str, subject: str, es_dict: dict ):
    """Sends email to the stakeholders"""

    ###### Get Sys_id from Database ##########
    query = f"select sys_id from incident_chat_map where chat_id='{chat_id}'"
    result = ENGINE.execute(query)
    sys_id = result.fetchone()['sys_id']

    if es_dict['Emistate'] in ["Declared", "In-Progress"]:
        th_color = '#FF0000'
    elif es_dict['Emistate'] in ["On-Hold", "Under Observation", "Restored/Monitoring"]:
        th_color = '#FFC000'
    elif es_dict['Emistate'] == "Resolved":
        th_color = "#00B050"
    else:
        th_color = "#00B0F0"

    #Get the Mail Template
    with open("mail_template.txt", "r", encoding='utf8') as file:
        html_mail_body = file.read()

    mail_template = Template(html_mail_body)

    mail_body = mail_template.substitute(
        th_color=th_color,
        Emistate=es_dict['Emistate'],
        Epriority=es_dict['Epriority'],
        EIncSummary=es_dict['EIncSummary'],
        EBizImp=es_dict['EBizImp'],
        EImpLoc=es_dict['EImpLoc'],
        EImpClient=es_dict['EImpClient'],
        EImpApp=es_dict['EImpApp'],
        EImpUsr=es_dict['EImpUsr'],
        EIsRepBy=es_dict['EIsRepBy'],
        EIncRef=es_dict['EIncRef'],
        EVendor=es_dict['EVendor'],
        EIncStart=es_dict['EIncStart'],
        EMimEng=es_dict['EMimEng'],
        EMIM=es_dict['EMIM'],
        ESupTeams=es_dict['ESupTeams'],
        EWrkArnd=es_dict['EWrkArnd'],
        EChange=es_dict['EChange'],
        ERFO=es_dict['ERFO'],
        EResTime=es_dict['EResTime'],
        EOutDur=es_dict['EOutDur'],
        ENxtUpd=es_dict['ENxtUpd'],
        updates=updates,
        EDL=es_dict['EDL']
    )

    ######## Creating and Sending Email using SNOW SMTP########
    #recipients = get_recipients()
    try:
        recipients = es_dict['EDL'].split(',')
        logger.debug("DLs -> %s :: %s", es_dict['EDL'], type(es_dict['EDL']))
    except:
        return "No DL"

    data = {
        "to": recipients,
        "subject": subject,
        "html": mail_body,
        "table_name": "incident",
        "table_record_id": sys_id,
    }

    response = requests.post(URL, auth=(USER, PWD), headers=HEADERS, data=json.dumps(data))

    # Check for HTTP codes other than 200
    if response.status_code != 200:
        logger.error(
            'Status: %s | '\
            'Headers: %s | '\
            'Error Response: %s',
            response.status_code, response.headers, response.json()
        )
        logger.info("Mail was not sent!")
        return None

    logger.info("Mail successfully sent")
    return "Success"

    # ######## Creating and Sending Email using Office365 SMTP ########
    # sender_email = 'iim_bot@outlook.com'
    # sender_name = 'IIM Bot'
    # password = 'Office@123'
    # s = smtplib.SMTP(host='smtp.office365.com', port=587)
    # s.starttls()
    # s.login(sender_email, password)

    # recipients = get_recipients()

    # mail = MIMEMultipart()
    # mail["From"] = formataddr((sender_name, sender_email))
    # mail["To"] = ', '.join(recipients)
    # mail["subject"] = subject

    # # Turn these into html MIMEText objects
    # html_MIME = MIMEText(mail_body, "html")

    # # Add HTML/plain-text parts to MIMEMultipart message
    # mail.attach(html_MIME)

    # # Create secure connection with server and send email
    # s.send_message(mail)

def send_mir(filename: str, EDL: str):
    """Sends mir to the stakeholders"""

    ######## Creating and Sending Email with MIR as attachment using Office365 SMTP ########
    smtp = smtplib.SMTP(host='smtp.office365.com', port=587)
    smtp.starttls()
    smtp.login(SENDER_ADDRESS, PASSWORD)

    msg = MIMEMultipart()
    msg['From'] = formataddr((SENDER_NAME, SENDER_ADDRESS))

    # recipients = ["bhakti.prabhu@iimbot.onmicrosoft.com"]
    try:
        recipients = EDL.split(',')
        logger.debug("DLs -> %s :: %s", EDL, type(EDL))
    except:
        return "No DL"

    msg["To"] = ', '.join(recipients)

    msg['Subject']="Major Incident Management Report"

    # add in the message body
    body = 'Hello All, <br><br> PFA the Major Incident Management Report. <br><br> Regard, <br> IIM Bot'
    msg.attach(MIMEText(body, 'html'))

    doc = MIMEApplication(open("Final-MIR.docx", 'rb').read())
    doc.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.attach(doc)

    # send the message via the server set up earlier.
    try:
        smtp.send_message(msg)
        del msg
    except:
        return None

    return "Success"

def get_recipients():
    """Gets recipients of email from database"""
    org = ['LTI-NAUT','LTI']
    if len(org) == 1:
        query = query = f"SELECT dl_email_id FROM mailing_list where organization = '{org[0]}'"
    else:
        query = f"SELECT dl_email_id FROM mailing_list where organization in {tuple(org)}"
    result = ENGINE.execute(query)
    if result.rowcount == 0:
        logger.error("Requested distribution list/s not found in the Database")
        return []
    recipients = []
    for mail_id in result.fetchall():
        recipients.append(mail_id[0])

    return recipients
    #return ["bhakti.prabhu@iimbot.onmicrosoft.com"]
