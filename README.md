# VHSBand Delivery Ticket Generator

This script generates mulch delivery tickets from an input delivery spreadsheet file.

This script uses `pipenv` so you'll need that installed. It's written and tested using Python 3.7.4 and should work with basically any version of Python 3.7.x

### GCP credentials.json file

You'll need a `credentials.json` file for running this. Go to the [Google Cloud Platform Console](https://console.cloud.google.com), login using admin credentials for vhsband.com, select an appropriate project and grab its OAuth2 credentials. It's important that the app type by "TV or Limited input device"

## Running

Download the delivery database spreadsheet as CSV, then run using
`python make_tickets.py -d <document id> -t <CSV file>`

where `<document id>` is the ID of the document into which tickets should be generated. e.g. 
```
https://docs.google.com/document/d/<DOC ID>/edit
```