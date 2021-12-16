# import smtplib
# from email.mime.text import MIMEText
# from email.mime.multipart import MIMEMultipart
# from email.utils import formataddr

import requests
import json
from string import Template

from actions.requester import ENGINE

def send_email_notification(chat_id: str, updates: str, subject: str, es_dict: dict ):

    ###### Get Sys_id from Database ##########
    q = "select sys_id from incident_chat_map where chat_id='{0}'".format(chat_id)
    result = ENGINE.execute(q)
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
    f = open("mail_template.txt", "r")
    html_mail_body = f.read()
    f.close()

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
    recipients = es_dict['EDL'].split(',')
    print('DL :: ', es_dict['EDL'], ' :: ', type(es_dict['EDL']))
    # subject = "MIM - "+es_dict['EImpClient']+" - "+es_dict['Epriority']+"| "+es_dict['Emistate']+" | <"+es_dict['EImpApp']+"> | "+es_dict['EIncRef']+" | "+es_dict['EIncSummary']
    
    # Set the request parameters
    url = 'https://dev60561.service-now.com/api/now/v1/email'
    user = 'admin'
    pwd = 'B@march1998'
    headers = {"Content-Type":"application/json","Accept":"application/json"}
    data = {
        "to": recipients,
        "subject": subject,
        "html": mail_body,
        "table_name": "incident",
        "table_record_id": sys_id,
    }

    response = requests.post(url, auth=(user, pwd), headers=headers, data=json.dumps(data))

    # Check for HTTP codes other than 200
    if response.status_code != 200: 
        print('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())
        print("\n Mail was not sent!")
        return None
    else:
        print("Mail successfully sent")
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
    # subject = "MIM - "+es_dict['EImpClient']+"-"+es_dict['Epriority']+"| "+es_dict['Emistate']+" | <"+es_dict['EImpApp']+"> | "+es_dict['EIncRef']+" | "+es_dict['EIncSummary']

    # # Turn these into html MIMEText objects
    # html_MIME = MIMEText(mail_body, "html")

    # # Add HTML/plain-text parts to MIMEMultipart message
    # mail.attach(html_MIME)

    # # Create secure connection with server and send email
    # s.send_message(mail)

    # #print(mail_body)


def get_recipients():
    return ["bhaktiprabhu98@gmail.com","bhakti.prabhu@lntinfotech.com"]
    #return ["bhaktiprabhu98@gmail.com"]


