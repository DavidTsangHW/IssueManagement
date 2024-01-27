import pyodbc
import pandas
import sqlalchemy
import urllib

import win32com
import win32com.client
import os
import datetime

from datetime import datetime

cnxn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:\\CKLS\\log.mdb;')
#cnxn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:\\CKLS\\log.mdb;')

print('Connected')

sql = "select  * from NumberSequence where tableName = 'Issue_Event' and fieldName = 'eventId'"

print(sql)

df = pandas.read_sql(sql,cnxn)
print(df['nextId'].iloc[0])

sql = "update NumberSequence set nextId = nextId + 1 where tableName = 'Issue_Event' and fieldName = 'eventId'"
cnxn.execute(sql)
cnxn.commit()

connection_string = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=.\\log.mdb;'
    r'ExtendedAnsiSQL=1;'
)
connection_uri = f"access+pyodbc:///?odbc_connect={urllib.parse.quote_plus(connection_string)}"
engine = sqlalchemy.create_engine(connection_uri)


def createIssue(message):

    sql = "select  * from NumberSequence where tableName = 'Issue' and fieldName = 'issueId'"

    df = pandas.read_sql(sql,cnxn)
    issueId = str(df['nextId'].iloc[0] +1)
    companyId = ''
    module = 'UNSPECIFIED'
    issueType = 'Unspecified'
    title = message.Subject    
    priority = 'Normal'
    status = 'Resolving'
    initiatedBy = 'David'
    createDateTime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    ls_issue = []
    ls_issue.append([issueId,companyId,module,issueType,title,priority,status,initiatedBy,createDateTime])

    sql = "select top  1 * from issue"
    df = pandas.read_sql(sql,cnxn)

    del df['VendorCaseId']
    del df['UpgradeIssue']
    del df['EndDateTime']

    ls_columns = df.columns.to_list()


    df_i = pandas.DataFrame(ls_issue,columns=ls_columns, dtype='object')

    print(df_i)

    df_i.to_sql('Issue', engine, if_exists='append', index=False)

    print(df_i)

    sql = "update NumberSequence set nextId = nextId + 1 where tableName = 'Issue' and fieldName = 'issueId'"
    cnxn.execute(sql)
    cnxn.commit()

    return

def createIssueEvent(message):

    sql = "select  * from NumberSequence where tableName = 'IssueEvent' and fieldName = 'eventId'"

    df = pandas.read_sql(sql,cnxn)
    issueId = str(df['nextId'].iloc[0] +1)
    companyId = ''
    module = 'UNSPECIFIED'
    issueType = 'Unspecified'
    title = message.Subject    
    priority = 'Normal'
    status = 'Resolving'
    initiatedBy = 'David'
    createDateTime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    ls_issue = []
    ls_issue.append([issueId,companyId,module,issueType,title,priority,status,initiatedBy,createDateTime])

    sql = "select top  1 * from issue"
    df = pandas.read_sql(sql,cnxn)

    del df['VendorCaseId']
    del df['UpgradeIssue']
    del df['EndDateTime']

    ls_columns = df.columns.to_list()


    df_i = pandas.DataFrame(ls_issue,columns=ls_columns, dtype='object')

    print(df_i)

    df_i.to_sql('Issue', engine, if_exists='append', index=False)

    print(df_i)

    sql = "update NumberSequence set nextId = nextId + 1 where tableName = 'Issue' and fieldName = 'issueId'"
    cnxn.execute(sql)
    cnxn.commit()

    return

def send_email(subject, body, recipients, attachments):

    outlookD =win32com.client.Dispatch("Outlook.Application")
    newMail = outlookD.CreateItem(0)
    newMail.Subject = subject

    recipients = sorted(recipients)

    newMail.To = ";".join(recipients)
    newMail.Body = body

    print(recipients)

    # attach files
    for attachment in attachments:
        print(attachment)
        newMail.Attachments.Add(attachment)
        
    newMail.Send()    

    return

outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

outlookD =win32com.client.Dispatch("Outlook.Application")

inbox=outlook.GetDefaultFolder(6)

messages=inbox.Items

message = messages.GetFirst()

ls_message = []


while message:
    
    try:
        print(message.senton.date())
    except:
        pass

    senderEmailAddress = ''
    try:
        senderEmailAddress = message.SenderEmailAddress
        print('Sender:', message.SenderEmailAddress)
    except:
        pass

    recipients = ''
    try:
        for recipient in message.Recipients:
            print(recipient)
            recipients = recipients + ',' + str(recipient)
            print('Recipient:', recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress)
    except:
        pass


    ls_message = []
    ls_message.append([message.EntryID,message.Subject,message.Body,senderEmailAddress,recipients])

    df_message = pandas.DataFrame(ls_message, columns=['EntryID','Subject', 'Body', 'SenderEmailAddress', 'Recipients'])

    print(df_message)

    df_message.to_sql('email', engine, if_exists='append', index=False)

    sql = "select * from [issue] where title = '" + message.Subject + "'"

    df = pandas.read_sql(sql,cnxn)

    if len(df) > 0:
        print(df)
    else:
        createIssue(message)


    message = messages.GetNext()


exit()

cnxn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:\\CKLS\\log.mdb;')
#cnxn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:\\CKLS\\log.mdb;')

print('Connected')

sql = "select top 10 * from issue"

print(sql)

df = pandas.read_sql(sql,cnxn)

#df.to_sql("issue2",cnxn,if_exists='replace',index=False)

print(str(len(df)) + ' records')

connection_string = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=.\\log.mdb;'
    r'ExtendedAnsiSQL=1;'
)
connection_uri = f"access+pyodbc:///?odbc_connect={urllib.parse.quote_plus(connection_string)}"
engine = sqlalchemy.create_engine(connection_uri)

df.to_sql('issue2', engine, if_exists='replace', index=False)


cnxn.close()

print('Done')
