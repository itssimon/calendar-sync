# calendar-sync
One-way synchronization of calendar events from Exchange server to Google Calendar

## What will synchronize?

* Events within the configured timeframe relative to today
* Event subject, location, date/time and attendees with valid email addresses
* The event body/description will not be synced! The description field in Google Calendar is instead used to store a hash.

## How to set up?

1. Create a new Google Calendar and take note of its Calendar ID (xxx@group.calendar.google.com), which can be found in the calendar settings under Calendar Address
2. Pull the code from this GitHub repository
3. Edit the `calendar_sync.ini` configuration file and follow the hints given within the file
4. Visit the [Google API Console](https://console.developers.google.com/) and ...
     1. Create a new project
     2. Under Library, enable the [Google Calendar API](https://console.developers.google.com/apis/api/calendar-json.googleapis.com/overview)
     3. Under Credentials, create a new OAuth client ID for application type Other and download it as a JSON file
     4. Rename the downloaded JSON file to `google_api_client_secret.json` and place it into the credentials folder (replace the pre-existing dummy file)
5. Install Python 3 and make sure `pip` is available
6. Install dependencies
     * Run `pip install git+https://github.com/linkedin/pyexchange`
     * Run `pip install --upgrade google-api-python-client`
6. Run the `calendar_sync.bat` by double-clicking on it (Windows) or run `python ./calendar_sync.py` in the command line. On first run a browser window will be opened and you're asked for authorization to manage your Google Calendar.

## Troubleshoot

Feel free to add an issue in GitHub and I am happy to help out.
