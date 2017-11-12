import re
import json
import base64
import hashlib
import logging
import configparser
import requests
import httplib2
import oauth2client
from datetime import datetime
from datetime import timedelta
from pytz import timezone
from apiclient import discovery
from pyexchange import Exchange2010Service, ExchangeNTLMAuthConnection


# Constants
RESPONSE_MAPPING = {
    'Unknown': 'needsAction',
    'Accepted': 'accepted',
    'TentativelyAccepted': 'tentative',
    'Declined': 'declined',
}

# Global variables
google_service = None

# Read configuration
config = configparser.ConfigParser()
config.read('calendar_sync.ini')

# Determine timeframe for which to get events
tz = timezone(config['General'].get('Timezone', 'UTC'))
start_date = datetime.now() - timedelta(days=abs(config['General'].getint('DaysPast', 7)))
end_date = datetime.now() + timedelta(days=abs(config['General'].getint('DaysFuture', 14)))


def get_events_from_exchange():
    conf = config['Exchange']
    verify_cert = conf.getboolean('VerifyCert', True)

    # Disable SSL warnings when certification verification is turned off
    if not verify_cert:
        requests.packages.urllib3.disable_warnings()

    decrypted_password = base64.b64decode(conf['Password']).decode('utf-8')
    connection = ExchangeNTLMAuthConnection(url=conf['URL'], username=conf['Username'], password=decrypted_password, verify_certificate=verify_cert)
    service = Exchange2010Service(connection)
    calendar = service.calendar()

    return calendar.list_events(start=start_date, end=end_date, details=True)


def init_google_calendar_service():
    global google_service

    if google_service is not None:
        return google_service

    conf = config['Google Calendar']
    store = oauth2client.file.Storage(conf['CredentialsFile'])
    credentials = store.get()

    if not credentials or credentials.invalid:
        flow = oauth2client.client.flow_from_clientsecrets(conf['ClientSecretFile'], 'https://www.googleapis.com/auth/calendar')
        flow.user_agent = 'Calendar Sync'
        credentials = oauth2client.tools.run_flow(flow, store)
        print('Storing credentials to', conf['CredentialsFile'])

    google_service = discovery.build('calendar', 'v3', http = credentials.authorize(httplib2.Http()))
    return google_service


def hash_event(event_body):
    event_json = json.dumps(event_body, sort_keys=True)
    return hashlib.sha1(event_json.encode('utf-8')).hexdigest()


def transform_event(event):
    # Add only attendees with valid email address (API doesn't accept other values)
    attendees = [
        {
            'displayName': person.name,
            'email': person.email,
            'responseStatus': RESPONSE_MAPPING.get(person.response, 'needsAction'),
        }
    for person in event.attendees if re.match(r"[^@]+@[^@]+\.[^@]+", person.email)]

    # Add Google Calendar address to attendees (API requires this)
    attendees.append(
        {
            'email': config['Google Calendar']['CalendarAddress'],
            'responseStatus': 'accepted',
        }
    )

    event_body = {
        'iCalUID': event.id,
        'summary': event.subject,
        'location': event.location,
        'start': {
            'dateTime': event.start.isoformat(),
            'timeZone': 'UTC',
        },
        'end': {
            'dateTime': event.end.isoformat(),
            'timeZone': 'UTC',
        },
        'organizer': {
            'displayName': event.organizer.name if event.organizer is not None else None,
            'email': event.organizer.email if event.organizer is not None else None,
        },
        'attendees': sorted(attendees, key=lambda k: k['email']),
        'reminders': {
            'useDefault': True,
        },
    }

    # Recognize all-day events
    if (event.end - event.start).total_seconds() % (24 * 60 * 60) == 0:
        event_body['start'] = {'date': event.start.astimezone(tz).date().isoformat()}
        event_body['end'] = {'date': (event.end - timedelta(seconds=1)).astimezone(tz).date().isoformat()}

    # Save hash of event in description field
    event_body['description'] = hash_event(event_body)

    return event_body


def get_events_from_google_calendar():
    init_google_calendar_service()
    events_result = google_service.events().list(calendarId=config['Google Calendar']['CalendarAddress'],
                        timeMin=start_date.isoformat() + 'Z',
                        timeMax=end_date.isoformat() + 'Z',
                        singleEvents=True, orderBy='startTime').execute()

    # Return events as dictionary for easy retrieval by iCalUID
    return {e.get('iCalUID'): e for e in events_result.get('items', []) if e.get('iCalUID')}


def main():
    # Initiate logger
    logger = logging.getLogger('calendar_sync')
    logger.setLevel(logging.DEBUG)
    logger_sh = logging.StreamHandler()
    logger_fh = logging.FileHandler('calendar_sync.log', 'a')
    formatter = logging.Formatter('%(asctime)s - %(message)s')
    logger_sh.setFormatter(formatter)
    logger_fh.setFormatter(formatter)
    logger.addHandler(logger_sh)
    logger.addHandler(logger_fh)

    # Get events from Exchange
    logger.info('Connecting to Exchange (takes a few seconds) ...')
    exchange_events = get_events_from_exchange()
    logger.info('Retrieved %d events from Exchange', exchange_events.count)

    # Get events from Google Calendar
    logger.info('Connecting to Google Calendar ...')
    google_events = get_events_from_google_calendar()
    logger.info('Retrieved %d events from Google Calendar', len(google_events))
    google_calendar_id = config['Google Calendar']['CalendarAddress']

    # Loop through Exchange events and add/update in Google Calendar as required
    for event in exchange_events.events:
        event_body = transform_event(event)

        if event_body['iCalUID'] in google_events:
            google_event = google_events[event_body['iCalUID']]
            if event_body['description'] != google_event.get('description', ''):
                google_service.events().update(calendarId=google_calendar_id, eventId=google_event['id'], body=event_body).execute()
                logger.debug('Updating event: %s', event_body['summary'])
            google_events.pop(event_body['iCalUID'])
        else:
            logger.debug('Importing new event: %s', event_body['summary'])
            google_service.events().import_(calendarId=google_calendar_id, body=event_body).execute()

    # Loop through remaining events in Google Calendar
    # These were not in the list of events retrieved from Exchange and will be deleted from Google Calendar
    for k, google_event in google_events.items():
        logger.debug('Deleting event: %s', google_event['summary'])
        google_service.events().delete(calendarId=google_calendar_id, eventId=google_event['id']).execute()


if __name__ == "__main__":
    main()
