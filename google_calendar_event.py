"""Modification of the url-text default script in OBS to get events from the Google Calendar API"""
# Created by burkdan@umich.edu

from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import obspython as obs
import urllib.request
import urllib.error

import datetime

cal_url     = ""
interval    = 30
source_name = ""
# source_names = {}

# ------------------------------------------------------------

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/calendar.readonly'
# CLIENT_SECRET_FILE = 'C:/users/Daniel/Desktop/School/Design/DS_projector/client_secret.json'
CLIENT_SECRET_FILE = ''
APPLICATION_NAME = 'Google Calendar OBS event script'

# Taken from the quickstart.py in the Google Calendar API documentation
def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'calendar-python.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run_flow(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def update_text():
    global cal_url
    global interval
    global source_name
    # global source_names
    global CLIENT_SECRET_FILE

    # if source is not None:
        
    # Gets stored credentials (taken from Calendar API quickstart)
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    # Time objects using datetime
    dt_now = datetime.datetime.utcnow()
    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    # Gets events currently happending by setting bounds to events happening within a second of current datetime, max ten events 
    events = service.events().list(calendarId=cal_url, timeMin=now, timeMax=(dt_now+datetime.timedelta(0,1)).isoformat() +'Z', maxResults=10, singleEvents=True,
                            orderBy='startTime').execute()
    
    for event in events['items']:
        print(event['summary'])


    text = events['items'][0]['summary']

    settings = obs.obs_data_create()
    obs.obs_data_set_string(settings, "text", text)
    source = obs.obs_get_source_by_name(source_name)
    obs.obs_source_update(source, settings)
    obs.obs_data_release(settings)
    obs.obs_source_release(source)

    # for event in events['items']:
    #     text = event['summary']
    #     settings = obs.obs_data_create()
    #     obs.obs_data_set_string(settings, "text", text)
    #     source = obs.obs_get_source_by_name(source_names[text])
    #     obs.obs_source_update(source, settings)
    #     obs.obs_data_release(settings)
    #     obs.obs_source_release(source)

def refresh_pressed(props, prop):
    update_text()

# ------------------------------------------------------------

def script_description():
    return "Upates text based on a Google Calendar event"

def script_update(settings):
    global cal_url
    global interval
    global source_name
    # global source_names
    global CLIENT_SECRET_FILE

    cal_url                = obs.obs_data_get_string(settings, "calendar_url")
    CLIENT_SECRET_FILE     = obs.obs_data_get_string(settings, "client_secret_file")
    interval               = obs.obs_data_get_int(settings, "interval")
    source_name            = obs.obs_data_get_string(settings, "source")
    # for source_name in source_names:
    #      obs.obs_data_get_string(settings, "source")
    

    obs.timer_remove(update_text)

    if cal_url != "": #and source_name != "":
        obs.timer_add(update_text, interval * 1000)

def script_defaults(settings):
    obs.obs_data_set_default_int(settings, "interval", 30)

def script_properties():
    props = obs.obs_properties_create()

    obs.obs_properties_add_text(props, "calendar_url", "Calendar URL", obs.OBS_TEXT_DEFAULT)
    obs.obs_properties_add_path(props, "client_secret_file", "Client Secret File", obs.OBS_PATH_FILE,'*.json', "")
    obs.obs_properties_add_int(props, "interval", "Update Interval (seconds)", 5, 3600, 1)

    p = obs.obs_properties_add_list(props, "source", "Text Source", obs.OBS_COMBO_TYPE_EDITABLE, obs.OBS_COMBO_FORMAT_STRING)
    sources = obs.obs_enum_sources()
    if sources is not None:
        for source in sources:
            source_id = obs.obs_source_get_id(source)
            if source_id == "text_gdiplus" or source_id == "text_ft2_source":
                name = obs.obs_source_get_name(source)
                obs.obs_property_list_add_string(p, name, name)

        obs.source_list_release(sources)

    obs.obs_properties_add_button(props, "button", "Refresh", refresh_pressed)
    return props
