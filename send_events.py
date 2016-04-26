from __future__ import print_function
import httplib2
import os

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools
import pytz
import datetime, dateutil.parser
import sqlite3 as lite
import time

import smtplib
import ConfigParser

SCOPES = 'https://www.googleapis.com/auth/calendar.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'

class ConfigError(Exception):
    pass

def parse_config(filename):
    """
    Parse input file into config dictionary. For each section in the input file it will create a key
    with the value being a dictionary of key value pairs of options from that section.
    """
    if not os.path.exists(filename):
        raise ConfigError('Config file does not exist.')
    parser = ConfigParser.RawConfigParser()
    parser.read(filename)
    cfg = {}
    for section in parser.sections():
        cfg[section] = dict(parser.items(section))
    return cfg
        

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
                                   'calendar-python-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
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

def get_events(cfg):
    """
    Returns events returned from API call.
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    calendar_url = cfg['calendar']['calendarid'] + '@group.calendar.google.com'
    eventsResult = service.events().list(
        calendarId=calendar_url, timeMin=now, singleEvents=True,
        orderBy='startTime', maxResults=40).execute()    # maxResults=40
    events = eventsResult.get('items', [])
    return events

def convert_date_hours(date_string):
    """
    Convert a date with specific hours (not a full day event) to an abbreviated format.
    Ex:
        6/4 7:00 PM
    """
    date = datetime.datetime.strptime(date_string[:19],"%Y-%m-%dT%H:%M:%S")

    #TODO: spend more time figuring out why this is 3 hours off.
    date = date + datetime.timedelta(hours=3)
    date = datetime.datetime.strftime(date, "%-m/%-d %-I:%M %p")
    return date

def convert_date_no_hours(date_string):
    """
    Convert an all day event date to the abbreviated format.
    Ex:
        6/4
    """
    date = datetime.datetime.strptime(date_string, "%Y-%m-%d")
    
    #TODO: spend more time figuring out why this is 3 hours off.
    date = date + datetime.timedelta(hours=3)
    date = datetime.datetime.strftime(date, "%-m/%-d") 
    return date

def parse_events(events):
    """
    Return a list of tuples of event data.
    """
    event_list = []
    for event in events:
        start = event['start'].get('dateTime')
        if start:
            start_date = convert_date_hours(start)
            end_date = None
        else:
            start = event['start'].get('date')
            start_date = convert_date_no_hours(start)
            
            end = event['end'].get('date')
            end_date = convert_date_no_hours(end)
 
        title = event['summary']
        if '@' not in title:
            # some titles follow the format "Event @ Location", others are just the Event name. 
            # If it's just the event name, get the location if it exists, and format it the same
            # way.
            if event.get('location'):
                location = event['location'].split(',', 1)[0]
                title = title + ' @ ' + location
        created = event['created']
        created_date = datetime.datetime.strptime(created[:19], "%Y-%m-%dT%H:%M:%S")
        
        # mark posts as new
        if (datetime.datetime.now() - created_date).days < 7:
            new = "NEW!"
        else:
            new = ""
        if end_date:
            date = start_date + " - " + end_date
        else:
            date = start_date
        data =  (date, title, new)
        event_list.append(data)
    return event_list


def generate_message(events):
    """
    Generate string of email message
    """
    #TODO: this should be made more broad so that someone could enter their own email.
    todays_date = datetime.datetime.now().strftime("%-m/%-d")
    
    event_string = "\n".join(["\t".join(i) for i in events])
    
    message = """From: Allston Rat Citizens
    \r\nSubject: Allston Rat City Weekly Digest ({})
    \r\n\r\nUpcoming events on the Allston Rat City Calendar:
    \n\n{}
    \n\nEvents marked as NEW! have been added to the calendar in the past week.""".format(todays_date, event_string)

    return message

def send_email(events, cfg):
    """
    Read email options from cfg and send email to recipients.
    """
    email_username = cfg['email']['username']
    email = email_username + '@gmail.com'
    FROM = email
    TO = cfg['email']['recipients'].split(',')

    password = cfg['email']['password']
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login(email_username, password)
    server.sendmail(FROM, TO, message)
    server.quit()
 

if __name__ == '__main__':
    import argparse
    argparser = argparse.ArgumentParser()
    argparser.add_argument('--no-email', action='store_true', default=False)
    args = argparser.parse_args()

    cfg = parse_config('calendar.cfg')
    events = get_events(cfg)
    if not events:
        print('No events found!')
    else:
        parsed_events = parse_events(events)
        message = generate_message(parsed_events)

        if args.no_email:
            print(message)
        else:
            send_email(message, cfg)
