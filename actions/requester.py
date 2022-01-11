from actions.auth_tokens import get_bot_headers
import requests
import json
import magic
from sqlalchemy import create_engine, inspect, MetaData, Table, Column, Integer, String
from sqlalchemy.engine.url import URL
import yaml

endpoint_file = open("endpoints.yml")
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
       Column('chat_id', String(100)),
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
USER = 'admin'
PWD = 'B@march1998'

def get_response(tableSpec: str, chat_id: str, queryFilter: str):

    ###### Get Sys_id from Database ##########
    q = "select sys_id from incident_chat_map where chat_id='{0}'".format(chat_id)
    result = ENGINE.execute(q)
    sys_id = result.fetchone()['sys_id']
    
    url = BASE_URL+tableSpec+sys_id+queryFilter

    # Set proper headers
    headers = {"Content-Type":"application/json","Accept":"application/json"}

    # Do the HTTP request
    response = requests.get(url, auth=(USER, PWD), headers=headers )

    # Check for HTTP codes other than 200
    if response.status_code != 200: 
        print('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())
        return None
    # Decode the JSON response into a dictionary and use the data
    data = response.json()
    return data['result']

def get_article(sys_id: str):
  
    url = BASE_URL+"kb_knowledge/"+sys_id+"?sysparm_fields=number,short_description,sys_view_count,use_count,text"

    # Set proper headers
    headers = {"Content-Type":"application/json","Accept":"application/json"}

    # Do the HTTP request
    response = requests.get(url, auth=(USER, PWD), headers=headers )

    # Check for HTTP codes other than 200
    if response.status_code != 200: 
        print('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())
        return None
    # Decode the JSON response into a dictionary and use the data
    data = response.json()
    return data['result']

def put_response(tableSpec: str, chat_id: str, queryFilter: str, data: json):

    ###### Get Sys_id from Database ##########
    q = "select sys_id from incident_chat_map where chat_id='{0}'".format(chat_id)
    result = ENGINE.execute(q)
    sys_id = result.fetchone()['sys_id']
    
    url = BASE_URL+tableSpec+sys_id+queryFilter

    # Set proper headers
    headers = {"Content-Type":"application/json","Accept":"application/json"}

    # Do the HTTP request
    response = requests.put(url, auth=(USER, PWD), headers=headers, data=data)
    data = response.json()

    if response.status_code == 403: 
        return response.status_code, data['error']

    # Check for HTTP codes other than 200
    if response.status_code != 200: 
        print('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())
        return None, None

    # Decode the JSON response into a dictionary and use the data
    return response.status_code, data['result']

def get_file_content(contentURL: str):

    headers = get_bot_headers()
    response = requests.get(contentURL, headers=headers)
    file_format = magic.from_buffer(response.content)
    if 'JPEG' in file_format:
        file_extension = 'jpg'
    elif 'PNG' in file_format:
        file_extension = 'png'
    else:
        file_extension = None
    return response.content, file_extension

def post_attachment(chat_id: str, file_name: str, data: str):
    
    ###### Get Sys_id from Database ##########
    q = "select sys_id from incident_chat_map where chat_Id='{0}'".format(chat_id)
    result = ENGINE.execute(q)
    sys_id = result.fetchone()['sys_id']

    url = 'https://dev60561.service-now.com/api/now/attachment/file?table_name=incident&table_sys_id='+sys_id+'&file_name='+file_name

    # Set proper headers
    headers = {"Content-Type":"image/*","Accept":"application/json"}

    # Do the HTTP request
    response = requests.post(url, auth=(USER, PWD), headers=headers, data=data)

    # Check for HTTP codes other than 201
    if response.status_code != 201: 
        print('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())
        return None

    # Decode the JSON response into a dictionary and use the data
    response_data = response.json()
    return response_data

def get_attachment(chat_id: str):
    
    ###### Get Sys_id from Database ##########
    q = "select sys_id from incident_chat_map where chat_id='{0}'".format(chat_id)
    result = ENGINE.execute(q)
    sys_id = result.fetchone()['sys_id']

    url = 'https://dev60561.service-now.com/api/now/attachment?sysparm_query=table_sys_id='+sys_id

    # Set proper headers
    headers = {"Content-Type":"application/json","Accept":"application/json"}

    # Do the HTTP request
    response = requests.get(url, auth=(USER, PWD), headers=headers)

    # Check for HTTP codes other than 200
    if response.status_code != 200: 
        print('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())
        return None

    # Decode the JSON response into a dictionary and use the data
    data = response.json()
    return data['result']

def get_groups():
    
    url = 'https://dev60561.service-now.com/api/now/table/sys_user_group?sysparm_fields=name'

    # Set proper headers
    headers = {"Content-Type":"application/json","Accept":"application/json"}

    # Do the HTTP request
    response = requests.get(url, auth=(USER, PWD), headers=headers )

    # Check for HTTP codes other than 200
    if response.status_code != 200: 
        print('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())
        return None

    # Decode the JSON response into a dictionary and use the data
    data = response.json()
    choiceList = []
    for group in data['result']:
        choice = {"title": group['name'],"value":group['name']}
        choiceList.append(choice)
    return choiceList