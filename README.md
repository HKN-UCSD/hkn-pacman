# hkn-pacman
Om Nom Nom ᗧ• • •

## Quickstart
https://developers.google.com/sheets/api/quickstart/python

## Prerequisites
- Python 2.6 or greater
- pip package management tool
- A Google Account
- Google Sheet Credentials (credentials.json)

## Install Setup
```
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib python-dotenv
```

## Add .env File
|                 Key                  |                Description                                   |
|--------------------------------------|--------------------------------------------------------------|
| SOURCE_EVENT_SHEET_ID                | Google Sheet ID of Event Point Log                           |
| SOURCE_EVENT_SHEET_RANGE             | Google Sheet Range of Event Point Log Data excluding headers |
| SOURCE_MENTOR_SHEET_ID               | Google Sheet ID of Mentor Point Log                          |
| SOURCE_MENTOR_SHEET_RANGE            | Google Sheet Range of Mentor Point Log Data exluding headers |
| DESTINATION_TOTAL_SHEET_ID           | Google Sheet ID of Total Point Log to be displayed           |
| DESTINATION_TOTAL_SHEET_RANGE_HEADER | Google Sheet Range of Title headers                          |
| DESTINATION_TOTAL_SHEET_RANGE_BODY   | Google Sheet Range of Total Point data                       |
| DESTINATION_DONE_SHEET_ID            | Google Sheet ID of Inductee emails who finished              |
| DESTINATION_DONE_SHEET_RANGE_BODY    | Google Sheet Range of Inductee emails who finished           |
## Cron Config
```
crontab -e 
0 22 * * * cd [hkn-pacman dir] && python pacman.py
```

## Progress
- [x] add total event points
- [x] add professional requirement check
- [x] add event list
- [x] add officer list
- [x] add mentor requirement check
- [x] add cron job (10pm daily)
- [x] clean code
- [x] write proper documentation

