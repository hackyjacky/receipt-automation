import getpass
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from pdf_annotate import PdfAnnotator, Location, Appearance
import smtplib, ssl
from email.message import EmailMessage
from googleapiclient.http import MediaFileUpload
import fitz

pw = getpass.getpass()
ct = int(input('ct: '))

tracker = 'trackerID'
doc = 'docID'
annex = 'annexID'
bundles = {'5':['Five', '1'], '15':['Fifteen', '2'],'25':['Twenty-Five', '3']}

smuCreds = None
myCreds = None

if os.path.exists('auth/smuToken.json'):
    smuCreds = Credentials.from_authorized_user_file('auth/smuToken.json', 'https://www.googleapis.com/auth/drive')

if not smuCreds or not smuCreds.valid:
    if smuCreds and smuCreds.expired and smuCreds.refresh_token:
        smuCreds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'auth/smuCreds.json', 'https://www.googleapis.com/auth/drive')
        smuCreds = flow.run_local_server(port=0)

    with open('auth/smuToken.json', 'w') as token:
        token.write(smuCreds.to_json())

if os.path.exists('auth/myToken.json'):
    myCreds = Credentials.from_authorized_user_file('auth/myToken.json', 'https://www.googleapis.com/auth/drive')

if not myCreds or not myCreds.valid:
    if myCreds and myCreds.expired and myCreds.refresh_token:
        myCreds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'auth/myCreds.json', 'https://www.googleapis.com/auth/drive')
        myCreds = flow.run_local_server(port=0)

    with open('auth/myToken.json', 'w') as token:
        token.write(myCreds.to_json())

sheetService = build('sheets', 'v4', credentials=smuCreds)
smuDriveService = build('drive', 'v3', credentials=smuCreds)
myDriveService = build('drive', 'v3', credentials=myCreds)
docService = build('docs', 'v1', credentials=smuCreds)

v = sheetService.spreadsheets().values()

while v.get(spreadsheetId=tracker, range=f'g{ct}').execute()['values'][0][0] == 'TRUE':
    row = v.get(spreadsheetId=tracker, range=f'c{ct}:e{ct}').execute()['values'][0]
    name = row[0]
    amt = row[2].split()[0][1:]
    bundle = bundles[amt]
    recno = f'clubNo-{ct - 1}'

    pdf = PdfAnnotator('r.pdf')
    pdf.add_annotation('text', Location(x1=670, y1=190, x2=780, y2=210, page=0), Appearance(content=recno, fill=(1, 0, 0), font_size=18))
    pdf.add_annotation('text', Location(x1=670, y1=160, x2=770, y2=180, page=0), Appearance(content='date', fill=(1, 0, 0), font_size=18))
    pdf.add_annotation('text', Location(x1=120, y1=140, x2=510, y2=160, page=0), Appearance(content=name, fill=(1, 0, 0), font_size=18))
    pdf.add_annotation('text', Location(x1=140, y1=110, x2=350, y2=140, page=0), Appearance(content=f'{bundle[0]} Dollars Only', fill=(1, 0, 0), font_size=18))
    pdf.add_annotation('text', Location(x1=140, y1=90, x2=400, y2=110, page=0), Appearance(content=f'Metamorphosis Camp Bundle {bundle[1]}', fill=(1, 0, 0), font_size=18))
    pdf.add_annotation('text', Location(x1=50, y1=50, x2=100, y2=80, page=0), Appearance(content=f'{amt}.00', fill=(1, 0, 0), font_size=18))
    pdf.add_annotation('text', Location(x1=660, y1=40, x2=760, y2=60, page=0), Appearance(content='issuerName', fill=(1, 0, 0), font_size=18))

    r = f'receipt{ct - 1}'
    pdf.write(f'pdf/{r}.pdf')
    
    msg = EmailMessage()
    msg['Subject'] = 'Receipt for purchase [Metamorphosis]'
    msg['From'] = 'issuerEmail'
    msg['To'] = row[1]
    msg.set_content('''Hi,

Thank you for signing up for Metamorphosis 2021! Attached is the receipt for your purchase. We hope you have fun during the camp!

Regards,
issuerName
Finance Executive
Singapore Management University | Metamorphosis
issuerEmail''')

    with open(f'pdf/{r}.pdf', 'rb') as pdf:
        msg.add_attachment(pdf.read(), maintype='application', subtype='pdf', filename=f'{r}.pdf')

    context = ssl.create_default_context()
    with smtplib.SMTP('smtp-mail.outlook.com', 587) as smtp:
        smtp.starttls(context=context)
        smtp.login('issuerEmail', pw)
        smtp.send_message(msg)
    
    smuDriveService.files().create(body={'name': f'{r}.pdf', 'parents': ['1Wpaa15_UGDsrmkFD_qmG5rpeyyTn912u']}, media_body=MediaFileUpload(f'pdf/{r}.pdf')).execute()

    with fitz.open(f'pdf/{r}.pdf') as pdf:
        img = pdf[0].get_pixmap()
        img.save(f'img/{r}.jpg')

    img = myDriveService.files().create(body={'name': f'{r}.jpg', 'parents': ['1vfwbb8l8hFj_VysSYxhAsre9oK0LPCUf']}, media_body=MediaFileUpload(f'img/{r}.jpg')).execute()

    os.remove(f'pdf/{r}.pdf')
    os.remove(f'img/{r}.jpg')

    idx = docService.documents().get(documentId=doc).execute()['body']['content'][-1]['endIndex']
    requests = [
        {
            'insertText': {
                'location': {
                    'index': idx - 1
                },
                'text': '\n\n'
            }
        },
        {
            'insertInlineImage': {
                'location': {
                    'index': idx - 1
                },
                'uri': f"https://docs.google.com/uc?id={img['id']}"
            }
        }
    ]
    docService.documents().batchUpdate(documentId=doc, body={'requests': requests}).execute()

    myDriveService.files().delete(fileId=img['id']).execute()

    sheetService.spreadsheets().values().update(spreadsheetId=tracker, range=f'h{ct}',valueInputOption='USER_ENTERED', body={'values': [['true']]}).execute()
    
    idx = (ct - 2) // 100
    remainder = (ct - 2) % 100
    if remainder == 0:
        id = sheetService.spreadsheets().get(spreadsheetId=annex, fields='sheets(properties(sheetId))').execute()['sheets'][-1]['properties']['sheetId']
        body = {
            'requests': {
                'duplicateSheet': {
                    'sourceSheetId': id, 'insertSheetIndex': idx, 'newSheetName': f'Statements of Accounts({idx + 1})'
                }
            }
        }
        sheetService.spreadsheets().batchUpdate(spreadsheetId=annex, body=body).execute()
        sheetService.spreadsheets().values().batchClear(spreadsheetId=annex, body={'ranges': [f'Statements of Accounts({idx + 1})!b12:61', f'Statements of Accounts({idx + 1})!g12:61']}).execute()

    if 0 <= remainder <= 49:
        row = f'b{remainder + 12}'
    else:
        row = f'g{remainder - 38}'

    sheetService.spreadsheets().values().update(spreadsheetId=annex, range=f'Statements of Accounts({idx + 1})!{row}',valueInputOption='USER_ENTERED', body={'values': [[name, recno, amt]]}).execute()

    ct += 1