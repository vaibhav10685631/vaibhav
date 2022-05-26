"""
This file contains all the constants.
"""

#Base query parameters for SNOW Tables
SLA_TABLE_SPEC = 'task_sla?sysparm_query=task.sys_idSTARTSWITH'
INC_TABLE_SPEC = 'incident/'
JRNL_TABLE_SPEC = 'sys_journal_field?sysparm_query=element_id='
KNLDG_TABLE_SPEC = 'm2m_kb_task?sysparm_query=task.sys_id='

#List of slots for Email communication
EMAIL_SLOTS = ['Epriority','Emistate','EIncSummary','EBizImp','EImpLoc','EImpClient',
        'EImpApp','EImpUsr','EIsRepBy','EIncRef','EVendor','EIncStart','EMimEng',
        'EMIM','ESupTeams','EWrkArnd','EChange','ERFO','EResTime','EOutDur','ENxtUpd','EDL']

#Base URL for BOT REST Communication
BOT_URL = "https://smba.trafficmanager.net/in/v3/conversations/"

#ID of APP added to Tenant App Catalog
CATALOG_APP_ID = "a776b851-d2ef-405d-8525-62e6b579fa9f"

#### Sharepoint Credentials ####
TENANT = 'iimbot'
SITE_NAME = 'MIR'
SITE_ID = '30caa8f1-56a6-4243-835f-4a270a97f1e0'
FOLDER = 'General'

#### IIM Bot Email Credentials ####
SENDER_ADDRESS = 'iim@iimbot.onmicrosoft.com'
SENDER_NAME = 'MIM'
PASSWORD = 'LTI@1234'

#### IIM ServiceNow Account Credentials ####
USER = 'iimbot'
PWD = 'B@march1998' 