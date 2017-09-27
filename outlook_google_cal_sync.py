from __future__ import print_function
from exchangelib import DELEGATE, IMPERSONATION, Account, Credentials, ServiceAccount, \
    EWSDateTime, EWSTimeZone, Configuration, NTLM, CalendarItem, Message, \
    Mailbox, Attendee, Q, ExtendedProperty, FileAttachment, ItemAttachment, \
    HTMLBody, Build, Version

from datetime import datetime
from datetime import timedelta
import dateutil.parser
import pytz
import filecmp
import shutil
import sqlite3
import re


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
    global new_cal_name, tz_name, outlook_user_name, outlook_user_pass, outlook_email, google_email, new_cal_id, max_fut_days_to_check, local_tz
    global datadir, SYNC_STATUS_NEW, SYNC_STATUS_EXIST_NOCHANGE, SYNC_STATUS_EXIST_CHANGED
    global bReinit

    bReinit = True # set to True if you want to reinitializes google calender by deleting it first.
    SYNC_STATUS_NEW = 1
    SYNC_STATUS_EXIST_NOCHANGE = 2
    SYNC_STATUS_EXIST_CHANGED = 3
    new_cal_name = ''
    tz_name = 'America/Chicago'
    outlook_user_name = ''
    outlook_user_pass = ''
    outlook_email = ''
    google_email = ''
    max_fut_days_to_check = 60
    local_tz = pytz.timezone(tz_name)
    datadir = "data"

    if bReinit:
        shutil.rmtree(datadir, ignore_errors=True)

    if not os.path.exists(datadir):
        os.makedirs(datadir)


    createDB()

def createDB():
    global db, datadir
    try:
        # Creates or opens a file called mydb with a SQLite3 DB
        db = sqlite3.connect(datadir + os.path.sep + "eventdb.sqlite")
        # Get a cursor object
        cursor = db.cursor()
        # Check if table users does not exist and create it
        cursor.execute('''CREATE TABLE IF NOT EXISTS
                          events(fileNo INTEGER PRIMARY KEY AUTOINCREMENT, outlook_event_id TEXT, CONSTRAINT u_evtid UNIQUE (outlook_event_id))''')

        # Commit the change
        db.commit()
    # Catch the exception
    except Exception as e:
        print("error : can not create Events DB")
        exit(0)
def runSQL(sql):
    global db
    cursor = db.cursor()
    cursor.execute(sql)
    db.commit()

def createEventFileMappingInDB(outlook_event_id):

    sql = "insert into events values (NULL," + "'" + str(outlook_event_id) + "')"
    runSQL(sql)

def getFileNoByOutlookEvtId(evtid):

    global db
    cursor = db.cursor()
    cursor.execute("SELECT fileNo FROM events where outlook_event_id='" + evtid + "'")
    row = cursor.fetchone()
    if row is None:
        return -1
    else:
        return row[0]

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

def cleanUpGoogleCal(calid):
    page_token = None
    while True:
      events = google_service.events().list(calendarId=calid, pageToken=page_token).execute()
      for event in events['items']:

         google_service.events().delete(calendarId=calid, eventId=event['id']).execute()
      page_token = events.get('nextPageToken')
      if not page_token:
        break


def createGoogleCal():
    global new_cal_id, bReinit
    # Delete all calenders with name new_cal_name
    calendar_list = google_service.calendarList().list().execute()
    bFound = False
    for cal in calendar_list['items']:
        if cal['summary'] == new_cal_name:
            bFound = True
            if bReinit:
                google_service.calendars().delete(calendarId=cal['id']).execute()
            new_cal_id = cal['id']


    if not bFound or bReinit:
        # create new calender with name new_cal_name
        calendar = {
            'summary': new_cal_name,
            'timeZone': tz_name
        }
        created_calendar = google_service.calendars().insert(body=calendar).execute()
        new_cal_id = created_calendar['id']

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

    return evt_id

def outlook_dt_to_google_dt(p_dt):

    dt = str(p_dt).replace(" ", "T")
    #google_dt = google_dt[0:19]
    google_dt = str(dt)
    #idx = dt.rfind(":")
    #google_dt = dt[:idx] + dt[idx+1:]
    return google_dt

def createBaseGoogleEventData(outlook_event):

    g_evt_id = outlook_evt_id_to_google_evt_id(outlook_event.item_id)
    start_dt = outlook_dt_to_google_dt(outlook_event.start.astimezone(local_tz))
    end_dt = outlook_dt_to_google_dt(outlook_event.end.astimezone(local_tz))




    google_event = {
          'id' : g_evt_id,
          'summary': outlook_event.subject + " : " + outlook_event.organizer.name ,
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
                    'timeZone': tz_name,
                    },
          'end': {
                    'dateTime': end_dt,
                    'timeZone': tz_name,
                  },
          'attendees': [
                        {'email': google_email},
                      ],
          'reminders': {
                        'useDefault': False,
                        'overrides': [
                            {'method': 'popup', 'minutes': 15},
                          ],
                        },
                      }
    if outlook_event.text_body != None:


        pattern = re.compile(r'http.*?webex.com.*?[ >\n]')
        webexarr = re.findall(pattern, outlook_event.text_body )
        if webexarr != None and len(webexarr) > 0:
            webexlink = webexarr[0]
            ch = webexlink[-1:]
            if ch == " " or ch == ">":
                webexlink = webexlink[:-1]

            google_event['description'] = webexlink


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

def isFutureDate(dt_check):

    if isinstance(dt_check, str):
            dt = dateutil.parser.parse(dt_check)
            timezone = dt.tzinfo
            if dt > datetime.now(timezone) + timedelta(hours=-2):
                return True
            else:
                return False

    tz = EWSTimeZone.timezone('GMT')
    t = datetime.now()
    dt = EWSDateTime(t.year, t.month, t.day, t.hour, t.minute, t.second)
    dt = tz.localize(dt)

    if dt_check > dt:
        return True
    else:
        return False

def handleExceptions(outlook_event, google_event):


    global gEvtList


    gEvtList = []


    bIsFutureEventModified = False
    bIsFutureEventDeleted = False
    if outlook_event.modified_occurrences == None and outlook_event.deleted_occurrences == None:
        return
    if outlook_event.modified_occurrences != None:
        for occ in outlook_event.modified_occurrences:
            if isFutureDate(occ.start):
                bIsFutureEventModified = True
    if outlook_event.deleted_occurrences != None:
        for occ in outlook_event.deleted_occurrences:
            if isFutureDate(occ.start):
                bIsFutureEventDeleted = True
    # create event list in google up to
    if bIsFutureEventModified or bIsFutureEventDeleted:
        page_token = None
        now = datetime.now()
        tz = EWSTimeZone.timezone('GMT')
        maxDt = str(datetime(year=now.year, month=now.month, day=now.day, tzinfo=tz) + timedelta(days=max_fut_days_to_check)).replace(" ", "T")

        while True:
            gevents = google_service.events().instances(calendarId=new_cal_id, eventId=google_event['id'],
                                                                pageToken=page_token, timeMax=maxDt).execute()
            for event in gevents['items']:
                    if isFutureDate(event['start']['dateTime']):
                            gEvtList.append(event)
            page_token = gevents.get('nextPageToken')
            if not page_token:
                    break

    if bIsFutureEventModified:
        handleModifiedOccurences(outlook_event)
    if bIsFutureEventDeleted:
        handleDeletedOccurences(outlook_event)

def findMatchingGoogleEventForRecurringInstance(occ):
    global gEvtList
    for event in gEvtList:

        if (str(type(occ)) == "<class 'exchangelib.recurrence.DeletedOccurrence'>"):
            occ_dt = occ.start
        else:
            occ_dt = occ.original_start

        google_dt = event['originalStartTime']['dateTime']
        idx = google_dt.rfind(":")
        google_dt = google_dt[:idx] + google_dt[idx+1:]
        dt = datetime.strptime(google_dt, '%Y-%m-%dT%H:%M:%S%z')
        if dt == occ_dt:
            return event
    return None

def handleModifiedOccurences(outlook_event):
    global gEvtList
    global bgoogle_event_edited

    for occ in outlook_event.modified_occurrences:
        if isFutureDate(occ.start):
            geventInstance = findMatchingGoogleEventForRecurringInstance(occ)

            if geventInstance != None:
                    outlook_event_instance = outlook_account.calendar.get(item_id=occ.item_id, changekey=occ.changekey)
                    if "STATUS:CANCELLED" in str(outlook_event_instance.mime_content):
                        google_service.events().delete(calendarId=new_cal_id, eventId=geventInstance['id']).execute()
                        bgoogle_event_edited = True

                    else:
                        geventInstance['start']['dateTime'] = str(occ.start).replace(" ", "T")
                        geventInstance['end']['dateTime'] = str(occ.end).replace(" ", "T")

                        updated_instance = google_service.events().update(calendarId=new_cal_id, eventId=geventInstance['id'], body=geventInstance).execute()
                        bgoogle_event_edited = True


def handleDeletedOccurences(outlook_event):
    global gEvtList
    for occ in outlook_event.deleted_occurrences:
        if isFutureDate(occ.start):
            geventInstance = findMatchingGoogleEventForRecurringInstance(occ)

            if geventInstance != None:
                    google_service.events().delete(calendarId=new_cal_id, eventId=geventInstance['id']).execute()

def removeTimeStampFromString(evtstr):
    return re.sub(r"DTSTAMP:.{17}", "DTSTAMP:", evtstr)

def checkPriorEventSyncStatus(outlook_event):
    global outlook_event_path, outlook_event_path_temp
    fileNo = getFileNoByOutlookEvtId(outlook_event.item_id)
    if fileNo == -1:
        return SYNC_STATUS_NEW
    outlook_event_path = datadir + os.path.sep + "outlook_event_" + str(fileNo)
    outlook_event_path_temp = outlook_event_path + ".temp"
    with open(outlook_event_path_temp, 'w') as file_temp:
        file_temp.write(removeTimeStampFromString(str(outlook_event)))
    if  filecmp.cmp(outlook_event_path, outlook_event_path_temp):
        return SYNC_STATUS_EXIST_NOCHANGE
    else:
        return SYNC_STATUS_EXIST_CHANGED


def isSingleEvent(outlook_event):
    if outlook_event.type != 'RecurringMaster':
        return True
    else:
        return False

def createRuleForWeeklyPattern(outlook_event):
        rule = []
        pattern = outlook_event.recurrence.pattern
        r = "RRULE:FREQ=WEEKLY"

        if str(type(outlook_event.recurrence.boundary)) == "<class 'exchangelib.recurrence.EndDatePattern'>":
            e = str(outlook_event.recurrence.boundary.end)[:10].replace("-", "")
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

        rule.append(r)
        return rule


def createGoogleEventData(outlook_event):

    event = createBaseGoogleEventData(outlook_event)
    if isSingleEvent(outlook_event):
        return event

    ### Handle Recurring Events
    if str(type(outlook_event.recurrence.pattern)) == "<class 'exchangelib.recurrence.WeeklyPattern'>":
        event['recurrence'] = createRuleForWeeklyPattern(outlook_event)

    return event



def sync_Events():
    global qs, bgoogle_event_edited, new_cal_id
    global outlook_event_path, outlook_event_path_temp

    qs = outlook_account.calendar.all()
    for outlook_event in qs:

        # if Event was synced before and there is no change then continue
        sync_status = checkPriorEventSyncStatus(outlook_event)

        if sync_status == SYNC_STATUS_EXIST_NOCHANGE:
            continue


        event = createGoogleEventData(outlook_event)
        if isSingleEvent(outlook_event):
            if not isFutureDate(outlook_event.start):
                continue
        if "STATUS:CANCELLED" in str(outlook_event.mime_content):
                if sync_status == SYNC_STATUS_EXIST_CHANGED:
                    google_service.events().delete(calendarId=new_cal_id, eventId=event['id']).execute()
                continue




        try:
            if sync_status == SYNC_STATUS_NEW:

                google_event = google_service.events().insert(calendarId=new_cal_id, body=event).execute()
                createEventFileMappingInDB(outlook_event.item_id)
                fileNo = getFileNoByOutlookEvtId(outlook_event.item_id)
                outlook_event_path = datadir + os.path.sep + "outlook_event_" + str(fileNo)

            elif sync_status == SYNC_STATUS_EXIST_CHANGED:
                google_event = google_service.events().update(calendarId=new_cal_id, eventId=event['id'], body=event).execute()

            if not isSingleEvent(outlook_event):
                bgoogle_event_edited = False
                handleExceptions(outlook_event, google_event)
                if bgoogle_event_edited:
                    google_event = google_service.events().get(calendarId=new_cal_id, eventId=google_event['id']).execute()

            fileNo = str(getFileNoByOutlookEvtId(outlook_event.item_id))
            google_event_path = datadir + os.path.sep + "google_event_" + fileNo

            if sync_status == SYNC_STATUS_EXIST_CHANGED:
                shutil.copy2(outlook_event_path, outlook_event_path + ".old")
                shutil.copy2(google_event_path, google_event_path + ".old")


            if sync_status == SYNC_STATUS_EXIST_CHANGED or sync_status == SYNC_STATUS_NEW:

                with open(google_event_path, 'w') as file:
                    file.write(removeTimeStampFromString(str(google_event)))
                with open(outlook_event_path, 'w') as file:
                    file.write(removeTimeStampFromString(str(outlook_event)))

        except HttpError as err:
            print("error occured in sync_Events : status - " + err.resp['status'] + " error text : " + str(err))




##### Main ######

init()
initOutlook()
initGoogle()
createGoogleCal()
sync_Events()
