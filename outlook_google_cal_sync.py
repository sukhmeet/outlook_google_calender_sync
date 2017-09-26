from __future__ import print_function
from exchangelib import DELEGATE, IMPERSONATION, Account, Credentials, ServiceAccount, \
    EWSDateTime, EWSTimeZone, Configuration, NTLM, CalendarItem, Message, \
    Mailbox, Attendee, Q, ExtendedProperty, FileAttachment, ItemAttachment, \
    HTMLBody, Build, Version

from datetime import datetime
from datetime import timedelta

import httplib2
from googleapiclient.errors import HttpError
import os
import string

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

def init():
    global new_cal_name, tz_name, outlook_user_name, outlook_user_pass, outlook_email, google_email, new_cal_id

    new_cal_name = ''
    tz_name = 'America/Chicago'
    outlook_user_name = ''
    outlook_user_pass = ''
    outlook_email = ''
    google_email = ''

def initOutlook():
    global outlook_account
    outlook_credentials = Credentials(username=outlook_user_name, password=outlook_user_pass)

    # Set up a target outlook_account and do an autodiscover lookup to find the target EWS endpoint:
    outlook_account = Account(primary_smtp_address=outlook_email, credentials=outlook_credentials,
                  autodiscover=True, access_type=DELEGATE)

    # If you're connecting to the same outlook_account very often, you can cache the autodiscover result for
    # later so you can skip the autodiscover lookup:
    ews_url = outlook_account.protocol.service_endpoint
    ews_auth_type = outlook_account.protocol.auth_type
    primary_smtp_address = outlook_account.primary_smtp_address

def initGoogle():
    global SCOPES, CLIENT_SECRET_FILE, APPLICATION_NAME, google_service
    SCOPES = 'https://www.googleapis.com/auth/calendar'
    CLIENT_SECRET_FILE = 'client_secret.json'
    APPLICATION_NAME = 'Google Calendar API Python Quickstart'
    google_credentials = get_GoogleCredentials()
    http = google_credentials.authorize(httplib2.Http())
    google_service = discovery.build('calendar', 'v3', http=http)

def get_GoogleCredentials():
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
                                   'calendar-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def recreateGoogleSyncCal():
    global new_cal_id
    # Delete all calenders with name new_cal_name
    calendar_list = google_service.calendarList().list().execute()

    for cal in calendar_list['items']:
        cal_id = cal['id']
        if cal['summary'] == new_cal_name:
            #print ("deleting calender " + new_cal_name)
            google_service.calendars().delete(calendarId=cal_id).execute()

#google_service.calendars().delete(cal_id).execute()

    # create new calender with name new_cal_name
    calendar = {
        'summary': new_cal_name,
        'timeZone': tz_name
    }
    created_calendar = google_service.calendars().insert(body=calendar).execute()
    new_cal_id = created_calendar['id']
    print(created_calendar['id'])

def outlook_evt_id_to_google_evt_id(p_id):
    evt_id = ""
    a_set = set(['w', 'x', 'y', 'z'])
    u_set = set ([])
    t = p_id.lower()
    for c in p_id:
        if c in string.ascii_lowercase and c not in a_set:
            evt_id += c
        elif c in a_set:
            if c == "w":
                evt_id += "0a"
            elif c == "x":
                evt_id += "1b"
            elif c == "y":
                evt_id += "2c"
            elif c == "z":
                evt_id += "3d"
        elif c in string.ascii_uppercase:
            evt_id += str(ord(c))
        elif c in string.digits:
            evt_id += c
    #print(p_id + " ### " + evt_id)
    return evt_id



def outlook_dt_to_google_dt(p_dt):

    google_dt = str(p_dt).replace(" ", "T")
    google_dt = google_dt[0:19]

    return google_dt

def getEvents():
    global qs
    qs = outlook_account.calendar.all()

def createBaseGoogleEventData(outlook_event):

    g_evt_id = outlook_evt_id_to_google_evt_id(outlook_event.item_id)
    start_dt = outlook_dt_to_google_dt(outlook_event.start)
    end_dt = outlook_dt_to_google_dt(outlook_event.end)

    google_event = {
          'id' : g_evt_id,
          'summary': outlook_event.organizer.name + " : " + outlook_event.subject ,
          'location': outlook_event.location,
          'description': '',
          'creator': {
                        'id': outlook_event.organizer.name,
                        'email': outlook_event.organizer.email_address,
                        'displayName': outlook_event.organizer.name,
                        'self': False
                     },
          'organizer': {
                        'id': outlook_event.organizer.name,
                        'email': outlook_event.organizer.email_address,
                        'displayName': outlook_event.organizer.name,
                        'self': False
                      },
          'start': {
                    'dateTime': start_dt,
                    'timeZone': 'GMT',
                    },
          'end': {
                    'dateTime': end_dt,
                    'timeZone': 'GMT',
                  },
          'attendees': [
                        {'email': google_email},
                      ],
          'reminders': {
                        'useDefault': False,
                        'overrides': [
                            {'method': 'popup', 'minutes': 15},
                            {'method': 'popup', 'minutes': 60},
                          ],
                        },
                      }
    return google_event

def outlook_wkday_to_google_wkday(day_num):
    if day_num == 1:
        return "MO"
    elif day_num == 2:
        return "TU"
    elif day_num == 3:
        return "WE"
    elif day_num == 4:
        return "TH"
    elif day_num == 5:
        return "FR"
    elif day_num == 6:
        return "SA"
    elif day_num == 7:
        return "SU"

def sync_Recurring_Events():
    global qs
    for item in qs:
        if item.type == 'RecurringMaster':
            event = createBaseGoogleEventData(item)

            pattern = item.recurrence.pattern

            if str(type(item.recurrence.pattern)) == "<class 'exchangelib.recurrence.WeeklyPattern'>":
                r = "RRULE:FREQ=WEEKLY"

                if str(type(item.recurrence.boundary)) == "<class 'exchangelib.recurrence.EndDatePattern'>":
                    e = str(item.recurrence.boundary.end)[:10].replace("-", "")
                    r += ";UNTIL="
                    r += e
                r += ";INTERVAL=" + str(pattern.interval)

                i=0;
                byday=''
                num = len(pattern.weekdays)
                for d in pattern.weekdays:
                    i = i+1
                    if i==1 and num > 1:
                        byday=";BYDAY=" + outlook_wkday_to_google_wkday(d) + ","
                    elif i==1 and num == 1:
                        byday=";BYDAY=" + outlook_wkday_to_google_wkday(d)
                    elif i < num:
                        byday=byday+ outlook_wkday_to_google_wkday(d) + ","
                    elif i == num:
                        byday=byday+ outlook_wkday_to_google_wkday(d)
                if num > 0:
                    byday+=";"
                r = r + byday
                recur = []
                recur.append(r)
                event['recurrence'] = recur

                #print("before creating recurring event : " + str(event))
                try:
                    eventRet = google_service.events().insert(calendarId=new_cal_id, body=event).execute()
                    #print("recurring event created " + event['start']['dateTime'] + event['summary'] + "\n" + str(event))
                except HttpError as err:
                    print("error occured :" + err.resp['status'])
                #exit(0)

            # s_str = "Pattern: Occurs on weekdays "
            # s_len = len(s_str) - 1
            # pattern = item.recurrence.pattern
            # if pattern[0:s_len] == s_str:
            #     print("After " + item.organizer , ":::" , item.start , ":::", item.end, ":::", item.subject , ":::", item.location, "\n\t", item.first_occurrence, "\n\t", item.last_occurrence, "\n\t", item.recurrence, "\n\t" ,item.modified_occurrences, "\n\t", item.deleted_occurrences )
            #     print("\n********n")
            #exit(0)

def sync_Single_Events():
    global qs
    tz = EWSTimeZone.timezone('GMT')
    t = datetime.now() + timedelta(days= - 5)
    dt = EWSDateTime(t.year, t.month, t.day)
    dt = tz.localize(dt)

    cal = qs.filter(start__gt=dt)
    for item in cal:
            if item.type != 'RecurringMaster':
                event = createBaseGoogleEventData(item)
                try:
                    #print("before creating single event : " + str(event))
                    eventRet = google_service.events().insert(calendarId=new_cal_id, body=event).execute()
                    #print("single event created " + event['start']['dateTime'] + event['summary'] + "\n" + str(event))
                except HttpError as err:
                    print("error occured :" + err.resp['status'])


##### Main ######
init()
initOutlook()
initGoogle()
recreateGoogleSyncCal()
getEvents()
sync_Single_Events()
sync_Recurring_Events()
