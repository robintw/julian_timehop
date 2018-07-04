from docx import Document
import datetime
import os
import pandas as pd
import emails
import dropbox


def read_table_from_doc(filename):
    doc = Document(filename)
    tab = doc.tables[0]

    dates = list(map(lambda x: x.text, tab.column_cells(0)))
    dates = pd.to_datetime(dates, errors='coerce')

    descs = list(map(lambda x: x.text, tab.column_cells(1)))

    df = pd.DataFrame({'desc':descs}, index=dates)
    
    return df

def get_formatted_message(df):
    today = datetime.datetime.now() - datetime.timedelta(2)
    today = today.strftime('%Y-%m-%d')
    
    subdf = df[today]
    
    if len(subdf) == 0:
        return
    
    title_date = datetime.datetime.now().strftime('%d %B')
    text = f'<h2>{title_date}</h2>\n\n'
    for i, row in subdf.iterrows():
        text += f"<p><b>{i.year!s}:</b></br>{row['desc']}</p>\n\n"
        
    return text

def send_email(text):
    message = emails.html(html=text,
                          subject='Julian Timehop',
                          mail_from=('Julian Timehop Emailer', 'robin@rtwilson.com'))

    password = os.environ.get('SMTP_PASSWORD')

    r = message.send(to=('R Wilson', 'robin@rtwilson.com'),
                     smtp={'host':'mail.rtwilson.com',
                           'port': 465, 'ssl': True,
                           'user': 'robin@rtwilson.com', 'password': password})

def download_file(path):
    dropbox_key = os.environ.get('DROPBOX_KEY')
    dbx = dropbox.Dropbox(dropbox_key)
    output_filename = 'document.docx'
    dbx.files_download_to_file(output_filename, path)
    
    return output_filename

filename = download_file('/Notes and diary entries for Julian.docx')
df = read_table_from_doc(filename)
text = get_formatted_message(df)
if text is not None:
    send_email(text)