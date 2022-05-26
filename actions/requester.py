"""
This file contains functions for fetching and pushing data to ServiceNow ITSM.
"""

import json
import logging
import yaml
import requests
import magic
from sqlalchemy import create_engine, inspect, MetaData, Table, Column, String
from sqlalchemy.engine.url import URL

from actions.auth_tokens import get_bot_headers
from actions.constants import USER, PWD

logger = logging.getLogger(__name__)

###### Create engine and check if table exists. If not, create new table. ######
with open("endpoints.yml", "r", encoding='utf8') as endpoint_file:
    parsed_endpoint_file = yaml.load(endpoint_file, Loader=yaml.FullLoader)
    tracker_store = parsed_endpoint_file["tracker_store"]

connect_url = URL.create(
    tracker_store['dialect'],
    username=tracker_store['username'],
    password=tracker_store['password'],
    host=tracker_store['url'],
    port=tracker_store['port'],
    database=tracker_store['db']
)

ENGINE = create_engine(connect_url)

insp = inspect(ENGINE)
if not insp.has_table("incident_chat_map"):
    meta = MetaData()

    incident_chat_map = Table(
       'incident_chat_map', meta,
       Column('number', String(30), primary_key = True),
       Column('sys_id', String(50)),
       Column('chat_id', String(100))
    )

    meta.create_all(ENGINE)

if not insp.has_table("email_updates"):
    meta = MetaData()

    email_updates = Table(
       'email_updates', meta,
       Column('number', String(30)),
       Column('date_time', String(35)),
       Column('update', String(500)),
       Column('manager', String(50))
    )

    meta.create_all(ENGINE)

BASE_URL = 'https://dev60561.service-now.com/api/now/table/'

def get_response(table_spec: str, chat_id: str, query_filter: str):
    """Fetches information from ServiceNow Tables"""

    ###### Get Sys_id from Database ##########
    query = f"select sys_id from incident_chat_map where chat_id='{chat_id}'"
    result = ENGINE.execute(query)
    sys_id = result.fetchone()['sys_id']

    url = BASE_URL+table_spec+sys_id+query_filter

    # Set proper headers
    headers = {"Content-Type":"application/json","Accept":"application/json"}

    # Do the HTTP request
    response = requests.get(url, auth=(USER, PWD), headers=headers )

    # Check for HTTP codes other than 200
    if response.status_code != 200:
        logger.error(
            'Status: %s | '\
            'Headers: %s | '\
            'Error Response: %s',
            response.status_code, response.headers, response.json()
        )
        return None
    # Decode the JSON response into a dictionary and use the data
    data = response.json()
    return data['result']

def get_article(sys_id: str):
    """Fetches knowledge article attached to a particular incident"""

    query_filter = "?sysparm_fields=number,short_description,sys_view_count,use_count,text"
    url = BASE_URL+"kb_knowledge/"+sys_id+query_filter

    # Set proper headers
    headers = {"Content-Type":"application/json","Accept":"application/json"}

    # Do the HTTP request
    response = requests.get(url, auth=(USER, PWD), headers=headers )

    # Check for HTTP codes other than 200
    if response.status_code != 200:
        logger.error(
            'Status: %s | '\
            'Headers: %s | '\
            'Error Response: %s',
            response.status_code, response.headers, response.json()
        )
        return None
    # Decode the JSON response into a dictionary and use the data
    data = response.json()
    return data['result']

def put_response(table_spec: str, chat_id: str, query_filter: str, data: json):
    """Pushes ticket data to ServiceNow ITSM"""

    ###### Get Sys_id from Database ##########
    query = f"select sys_id from incident_chat_map where chat_id='{chat_id}'"
    result = ENGINE.execute(query)
    sys_id = result.fetchone()['sys_id']

    url = BASE_URL+table_spec+sys_id+query_filter

    # Set proper headers
    headers = {"Content-Type":"application/json","Accept":"application/json"}

    # Do the HTTP request
    response = requests.put(url, auth=(USER, PWD), headers=headers, data=data)
    data = response.json()

    if response.status_code == 403:
        return response.status_code

    # Check for HTTP codes other than 200
    if response.status_code != 200:
        logger.error(
            'Status: %s | '\
            'Headers: %s | '\
            'Error Response: %s',
            response.status_code, response.headers, response.json()
        )
        return None

    # Decode the JSON response into a dictionary and use the data
    return response.status_code

def get_file_content(content_url: str):
    """Gets binary data of file attached in the channel"""

    headers = get_bot_headers()
    response = requests.get(content_url, headers=headers)
    file_format = magic.from_buffer(response.content)
    if 'JPEG' in file_format:
        file_extension = 'jpg'
    elif 'PNG' in file_format:
        file_extension = 'png'
    else:
        file_extension = None
    return response.content, file_extension

def post_attachment(chat_id: str, file_name: str, data: str):
    """Attaches the file to particular incident in ITSM"""

    ###### Get Sys_id from Database ##########
    query = f"select sys_id from incident_chat_map where chat_Id='{chat_id}'"
    result = ENGINE.execute(query)
    sys_id = result.fetchone()['sys_id']
    attachment_url = BASE_URL.replace('table','attachment')
    url = attachment_url+'file?table_name=incident&table_sys_id='+sys_id+'&file_name='+file_name

    # Set proper headers
    headers = {"Content-Type":"image/*","Accept":"application/json"}

    # Do the HTTP request
    response = requests.post(url, auth=(USER, PWD), headers=headers, data=data)

    # Check for HTTP codes other than 201
    if response.status_code != 201:
        logger.error(
            'Status: %s | '\
            'Headers: %s | '\
            'Error Response: %s',
            response.status_code, response.headers, response.json()
        )
        return None

    return response.status_code

def get_attachment(chat_id: str):
    """Fetches attachments for a particular incident from ITSM"""

    ###### Get Sys_id from Database ##########
    query = f"select sys_id from incident_chat_map where chat_id='{chat_id}'"
    result = ENGINE.execute(query)
    sys_id = result.fetchone()['sys_id']

    url = 'https://dev60561.service-now.com/api/now/attachment?sysparm_query=table_sys_id='+sys_id

    # Set proper headers
    headers = {"Content-Type":"application/json","Accept":"application/json"}

    # Do the HTTP request
    response = requests.get(url, auth=(USER, PWD), headers=headers)

    # Check for HTTP codes other than 200
    if response.status_code != 200:
        logger.error(
            'Status: %s | '\
            'Headers: %s | '\
            'Error Response: %s',
            response.status_code, response.headers, response.json()
        )
        return None

    # Decode the JSON response into a dictionary and use the data
    data = response.json()
    return data['result']

def get_groups():
    """Fetches list of Assignment Groups from ITSM"""

    url = 'https://dev60561.service-now.com/api/now/table/sys_user_group?sysparm_fields=name'

    # Set proper headers
    headers = {"Content-Type":"application/json","Accept":"application/json"}

    # Do the HTTP request
    response = requests.get(url, auth=(USER, PWD), headers=headers )

    # Check for HTTP codes other than 200
    if response.status_code != 200:
        logger.error(
            'Status: %s | '\
            'Headers: %s | '\
            'Error Response: %s',
            response.status_code, response.headers, response.json()
        )
        return None

    # Decode the JSON response into a dictionary and use the data
    data = response.json()
    choice_list = []
    for group in data['result']:
        choice = {"title": group['name'],"value":group['name']}
        choice_list.append(choice)
    return choice_list
