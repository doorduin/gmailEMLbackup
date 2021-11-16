# gmailEMLbackup
Simple Script to backup Emails to EML format from GMAIL

Over the years my GMAIL has reached its maximum capacity. This is a simple script to backup locally using the GMAIL api: https://developers.google.com/gmail/api

Ensure the following is installed via pip:
```
pip install --upgrade google-api-python-client oauth2client
pip install --humanfriendly
```

Example usage to get going quickly:
```
python gmailEMLbackup.py --year 2000 --subject --purging
```

To get the most up to date help and usage execute the following:
```
python gmailEMLbackup.py -h
```

The first time the script is run it will authenticate to the Google API's. The app accessing the details is registered as: gmail EML Backup on https://console.cloud.google.com/
