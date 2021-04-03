from __future__ import print_function
import pickle
import os.path
import argparse
import sys
import json
import logging
import csv
import time
from functools import cmp_to_key
from pprint import pformat
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.auth.transport.requests import AuthorizedSession

from consts import *

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/documents']
DISCOVERY_DOC = ('https://docs.googleapis.com/$discovery/rest?'
                 'version=v1')
# 2019 full tickets
# DOC_ID = '1OW0mKpL9tbTBbB3hcjo4GhAF5NKfagSCqF-2pKCTXVQ'

# test tickets
#DOC_ID = '1w-nGqeiom5k8SC3LJLFSjrUZPouha5gqhuSYBvfqW0g'

# test generated tickets
DOC_ID = '1otlM_JIfaywU5l5TwwgYb6snMp5dei1vxeNdeTIt15M'


def args_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument('-d', '--document', help='The ID of the document into which tickets are to be generated', type=str, default=DOC_ID)
    parser.add_argument('-t', '--tickets',
                        help='The filename of the CSV containing the final delivery tickets, extracted from Google Delivery database',
                        type=str, required=True)
    return parser

def parse_csv(filename):
    rows = []
    with open(filename, mode='r') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            rows.append(row)            
    return rows

def load_credentials():
    log = logging.getLogger('make_tickets')
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_console()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds

def build_google_api(service, version):
    log = logging.getLogger('make_tickets')

    log.info('getting credentials')
    creds = load_credentials()   
    log.info('building service object') 
    svc = build(service, version, credentials=creds, discoveryServiceUrl=DISCOVERY_DOC)
    return svc

def insert_ticket(ticket_data, target_doc, service):
    log = logging.getLogger('make_tickets')

    log.debug(pformat(ticket_data))

    txtStyBR = TEXT_STYLE_BOTTOMRIGHT
    txtStyBL = TEXT_STYLE_BOTTOMLEFT
    txtStyMR = TEXT_STYLE_MIDRIGHT
    txtStyML = TEXT_STYLE_MIDLEFT
    txtStyTR = TEXT_STYLE_TOPRIGHT
    txtStyTL = TEXT_STYLE_TOPLEFT

    pSty = PARAGRAPH_STYLE_MID_LEFT

    stpText = ''
    if ticket_data[STEEP_DRIVEWAY] == 'Checked':
        stpText = 'Steep'
    else:
        stpText = ' '
    txtStyBR['range']['endIndex'] = txtStyBR['range']['startIndex'] + len(stpText)
    delivInstr = ticket_data[DELIVERY_INSTR]
    if (delivInstr == ''):
        delivInstr = ' '
    if (len(delivInstr) > 60):
        txtStyBL = TEXT_STYLE_BOTTOMLEFT_SMALL
    txtStyBL['range']['endIndex'] = txtStyBL['range']['startIndex'] + len(delivInstr)
    delivZone = ticket_data[DELIVERY_ZONE]
    if delivZone == '':
        delivZone = 'UNKNOWN'
    txtStyMR['range']['endIndex'] = txtStyMR['range']['startIndex'] + len(delivZone)
    contact = ticket_data[FIRST_NAME] + ' ' + ticket_data[LAST_NAME] + ' ' + ticket_data[PHONE]    
    txtStyML['range']['endIndex'] = txtStyML['range']['startIndex'] + len(contact)
    pSty['range']['endIndex'] = pSty['range']['startIndex'] + len(contact)
    order = ticket_data[TOTAL_BAGS] + ' ' + ticket_data[MULCH_TYPE] + '\n#' + ticket_data[ORDER_NUMBER]
    if 'Hardwood' == ticket_data[MULCH_TYPE] or int(ticket_data[TOTAL_BAGS]) > 99:
        txtStyTR = TEXT_STYLE_TOPRIGHT_SMALL
    txtStyTR['range']['endIndex'] = txtStyTR['range']['startIndex'] + len(order)
    addr = ticket_data[ADDRESS]
    if ticket_data[GATE_CODE] != '':
        addr = "{}\nGate Code: {}".format(ticket_data[ADDRESS], ticket_data[GATE_CODE])
        txtStyTL = TEXT_STYLE_TOPLEFT_SMALL
    txtStyTL['range']['endIndex'] = txtStyTL['range']['startIndex'] + len(addr)

    requests = [
    { 'insertTable': { 'rows': 3, 'columns': 2, 'location': { 'index': 1 } } },
    { 'updateTableCellStyle': CELL_STYLE_TOPLEFT },
    { 'updateTableCellStyle': CELL_STYLE_TOPRIGHT },
    { 'updateTableCellStyle': CELL_STYLE_MIDLEFT },
    { 'updateTableCellStyle': CELL_STYLE_MIDRIGHT },        
    { 'updateTableCellStyle': CELL_STYLE_BOTTOMLEFT },
    { 'updateTableCellStyle': CELL_STYLE_BOTTOMRIGHT },    
    { 'insertText': { 'text': stpText, 'location': { 'index': 17 } } },    
    { 'updateTextStyle': txtStyBR },
    { 'insertText': { 'text': delivInstr, 'location': { 'index': 15 } } },
    { 'updateTextStyle': txtStyBL },
    { 'insertText': { 'text': delivZone, 'location': { 'index': 12 } } },
    { 'updateTextStyle': txtStyMR },
    { 'insertText': { 'text': contact, 'location': { 'index': 10 } } },
    { 'updateParagraphStyle': pSty },
    { 'updateTextStyle': txtStyML },
    { 'insertText': { 'text': order, 'location': { 'index': 7 } } },
    { 'updateTextStyle': txtStyTR },
    { 'insertText': { 'text': addr, 'location': { 'index': 5 } } },
    { 'updateTextStyle': txtStyTL },
    { 'updateTableRowStyle': ROW_STYLE_TALL },
    { 'updateTableRowStyle': ROW_STYLE_SHORT },
    { 'updateTableColumnProperties': COLUMN_PROPERTIES_WIDE },
    { 'updateTableColumnProperties': COLUMN_PROPERTIES_THIN }
    ]

    return requests
    # log.debug(pformat(requests))
    # log.info('Sending request')
    # result = service.documents().batchUpdate(documentId=target_doc, body={'requests': requests}).execute()
    # log.debug(pformat(result))

def row_cmp(item1, item2):
    log = logging.getLogger('row_cmp')
    log.debug(pformat(item1))
    log.debug(pformat(item2))
    zip1 = item1[ZIP]
    zip2 = item2[ZIP]
    bag1 = item1[TOTAL_BAGS]
    bag2 = item2[TOTAL_BAGS]
    zone1 = item1[DELIVERY_ZONE]
    zone2 = item2[DELIVERY_ZONE]

    if (bag1 == ''):
        bag1 = '0'
    if (bag2 == ''):
        bag2 = '0'
    if (int(bag1) > 51 and int(bag2) < 52):
        return -1
    elif (int(bag1) < 52 and int(bag2) > 51):
        return 1

    if (zip1 == '78732' and zip2 != '78732'):
        return -1
    elif (zip2 == '78732' and zip1 != '78732'):
        return 1

    if (zone1 == sorted([zone1,zone2])[0]):
        return -1
    else:
        return 1

def sort_rows(rows):
    log = logging.getLogger('sort_rows')
    new_rows = sorted(rows, key=cmp_to_key(row_cmp))
    log.debug('sorted_rows')
    log.debug(pformat(new_rows))
    return new_rows

def make_tickets(rows, target_doc):
    log = logging.getLogger('make_tickets')
    log.debug('getting service')
    docs = build_google_api('docs', 'v1')

    log.info('Creating tickets')
    idx = 0
    requests = []
    pageBreakReq = { 'insertPageBreak': { 'location': { 'index': 1 } } } 
    counter = 1

    for r in rows:
        log.info('Creating for ' + r[ORDER_NUMBER])
        requests.append(insert_ticket(r, target_doc, docs))        
        idx = idx + 1

        if idx % 4 == 0:
            requests.append(pageBreakReq)

        log.info('Sending requests')
        log.debug(pformat(requests))
        result = docs.documents().batchUpdate(documentId=target_doc, body={'requests': requests}).execute()
        log.debug(pformat(result))
        requests = []
        log.info("{} / {}".format(counter, len(rows)))
        counter = counter + 1
        log.info('Sleeping for 1 second to avoid quota')
        time.sleep(1)


def main():
    logging.basicConfig(level=logging.INFO)
    logging.getLogger('main').setLevel(logging.DEBUG)
    logging.getLogger('googleapiclient').setLevel(logging.ERROR)
    log = logging.getLogger('main')
    parser = args_parser()
    args = parser.parse_args()
    rows = parse_csv(args.tickets)
    rows = sort_rows(rows)
    # for row in rows:
    #     log.debug("{},{},{}".format(row[ZIP], row[TOTAL_BAGS], row[DELIVERY_ZONE]))
    make_tickets(rows, args.document)

#test code

#Test function to dump the JSON contents of a given document
def dump_json():
    docs = build_google_api('docs', 'v1')

    result = docs.documents().get(documentId=DOC_ID).execute()
    print(json.dumps(result, indent=4, sort_keys=True))

def run_test():
    log = logging.getLogger('make_tickets')

    log.debug('getting service')
    docs = build_google_api('docs', 'v1')

    test_data = {
        STEEP_DRIVEWAY: True,
        DELIVERY_INSTR: 'Place by the top right of the driveway. thanks',
        DELIVERY_ZONE: 'Shire Ridge',
        FIRST_NAME: 'John',
        LAST_NAME: 'Delancy',
        PHONE: '512-555-1414',
        TOTAL_BAGS: 15,
        MULCH_TYPE: 'Black',
        ORDER_NUMBER: '40-1111',
        ADDRESS: '123 Fake Ave'
    }

    log.info('Creating tickets')
    insert_ticket(test_data, docs)

def static_test():
    log = logging.getLogger('make_tickets')

    log.debug('getting service')
    docs = build_google_api('docs', 'v1')

    requests = [
    { 'insertTable': { 'rows': 3, 'columns': 2, 'location': { 'index': 1 } } },
    { 'updateTableCellStyle': CELL_STYLE_TOPLEFT },
    { 'updateTableCellStyle': CELL_STYLE_TOPRIGHT },
    { 'updateTableCellStyle': CELL_STYLE_MIDLEFT },
    { 'updateTableCellStyle': CELL_STYLE_MIDRIGHT },        
    { 'updateTableCellStyle': CELL_STYLE_BOTTOMLEFT },
    { 'updateTableCellStyle': CELL_STYLE_BOTTOMRIGHT },    
    { 'insertText': { 'text': 'Steep Driveway', 'location': { 'index': 17 } } },    
    { 'updateTextStyle': { 'range': { 'startIndex': 17, 'endIndex': 31 }, 'textStyle': { 'bold': 'true', 'weightedFontFamily': { 'fontFamily': 'Times New Roman' }, 'fontSize': { 'magnitude': 16, 'unit': 'PT' } }, 'fields': '*' } },
    { 'insertText': { 'text': 'Some delivery instructions go here', 'location': { 'index': 15 } } },
    { 'updateTextStyle': { 'range': { 'startIndex': 15, 'endIndex': 15+len('Some delivery instructions go here') }, 'textStyle': { 'weightedFontFamily': { 'fontFamily': 'Times New Roman' }, 'fontSize': { 'magnitude': 14, 'unit': 'PT' } }, 'fields': '*' } },
    { 'insertText': { 'text': 'River Ridge', 'location': { 'index': 12 } } },
    { 'updateTextStyle': { 'range': { 'startIndex': 12, 'endIndex': 12+len('River Ridge') }, 'textStyle': { 'weightedFontFamily': { 'fontFamily': 'Times New Roman' }, 'fontSize': { 'magnitude': 20, 'unit': 'PT' } }, 'fields': '*' } },
    { 'insertText': { 'text': 'Troy Mclure 512-555-1212', 'location': { 'index': 10 } } },
    { 'updateParagraphStyle': { 'range': { 'startIndex': 10, 'endIndex': 10+len('Troy Mclure 512-555-1212') }, 'paragraphStyle': { 'namedStyleType': 'NORMAL_TEXT', 'alignment': 'END' }, 'fields': '*' } },
    { 'updateTextStyle': { 'range': { 'startIndex': 10, 'endIndex': 10+len('Troy Mclure 512-555-1212') }, 'textStyle': { 'weightedFontFamily': { 'fontFamily': 'Times New Roman' }, 'fontSize': { 'magnitude': 18, 'unit': 'PT' } }, 'fields': '*' } },
    { 'insertText': { 'text': '52 HardW\n#30-1111', 'location': { 'index': 7 } } },
    { 'updateTextStyle': { 'range': { 'startIndex': 7, 'endIndex': 7+len('52 HardW\n#30-1111') }, 'textStyle': { 'bold': 'true', 'weightedFontFamily': { 'fontFamily': 'Times New Roman' }, 'fontSize': { 'magnitude': 20, 'unit': 'PT' } }, 'fields': '*' } },
    { 'insertText': { 'text': '1234 Fake Street', 'location': { 'index': 5 } } },
    { 'updateTextStyle': { 'range': { 'startIndex': 5, 'endIndex': 5+len('1234 Fake Street') }, 'textStyle': { 'weightedFontFamily': { 'fontFamily': 'Times New Roman' }, 'fontSize': { 'magnitude': 20, 'unit': 'PT' } }, 'fields': '*' } },
    { 'updateTableRowStyle': ROW_STYLE_TALL },
    { 'updateTableRowStyle': ROW_STYLE_SHORT },
    { 'updateTableColumnProperties': COLUMN_PROPERTIES_WIDE },
    { 'updateTableColumnProperties': COLUMN_PROPERTIES_THIN }
    ]

    result = docs.documents().batchUpdate(documentId=DOC_ID, body={'requests': requests}).execute()
    log.info(pformat(result))


if __name__ == '__main__':
    main()
