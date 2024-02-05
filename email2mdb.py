import pyodbc
import pandas
import sqlalchemy
import urllib

import win32com
import win32com.client
import os
import datetime

from datetime import datetime


server = 'TESTAX4'
database = 'DataTools'

connstr = 'DRIVER={SQL Server Native Client 11.0};SERVER='+server+';DATABASE='+database+';Trusted_Connection=yes;'
conn2= urllib.parse.quote_plus(connstr)
engine2 = sqlalchemy.create_engine('mssql+pyodbc:///?odbc_connect={}'.format(conn2))
cnxn2 = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER='+server+';DATABASE='+database+';Trusted_Connection=yes')

print('SQL SERVER CONNECTED')


cnxn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:\\CKLS\\log.mdb;')
#cnxn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:\\CKLS\\log.mdb;')

print('ACCESS DB CONNECTED')

sql = "select * from NumberSequence where tableName = 'Issue_Event' and fieldName = 'eventId'"

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


def createIssue(EntryID):

    sql = "select  * from NumberSequence where tableName = 'Issue' and fieldName = 'issueId'"

    df = pandas.read_sql(sql,cnxn)

    seq = df.iloc[0]

    recipients = ''
    sql = "select * from TBL_DLOG_EMAIL where EntryID = '" + EntryID + "'"
    df_e = pandas.read_sql(sql,cnxn2)
    email = df_e.iloc[0]
        
    issueId = str(seq['nextId']).rjust(seq['sequenceLength']+1, seq['leadingCharacter'])

    companyId = ''
    module = 'UNSPECIFIED'
    issueType = 'Unspecified'
    title = email['Subject']
    priority = 'Normal'
    if message.Importance == 2:
        priority = 'Higher'
    if message.Importance == 0:
        priority = 'Lower'
    
    status = 'Resolving'

    initiatedBy = email['Recipients']

    if len(initiatedBy) == 0:
        initiatedBy = 'Unknown'

    createDateTime = email['ReceivedTime']

    ls_issue = []
    ls_issue.append([issueId,companyId,module,issueType,title,priority,status,initiatedBy,createDateTime])
    
    sql = "select top  1 * from issue"
    df = pandas.read_sql(sql,cnxn)

    del df['VendorCaseId']
    del df['EndDateTime']

    ls_columns = df.columns.to_list()

    df_i = pandas.DataFrame(ls_issue,columns=ls_columns, dtype='object')
    
    df_i.to_sql('Issue', engine, if_exists='append', index=False)

    sql = "update NumberSequence set nextId = nextId + 1 where tableName = 'Issue' and fieldName = 'issueId'"
    cnxn.execute(sql)
    cnxn.commit()

    return df_i.iloc[0]

def createIssueEvent(EntryID, issue):

    sql = "select  * from NumberSequence where tableName = 'Issue_Event' and fieldName = 'eventId'"

    df = pandas.read_sql(sql,cnxn)

    sql = "select * from TBL_DLOG_EMAIL where EntryID = '" + EntryID + "'"
    df_e = pandas.read_sql(sql,cnxn2)

    email = df_e.iloc[0]    

    seq  = df.iloc[0]

    eventId = str(seq['nextId']).rjust(seq['sequenceLength']+1, seq['leadingCharacter'])

    issueId = issue['IssueID']
    userName = email['Recipients']
    description = email['Body'].strip()
    recordDateTime =  email['ReceivedTime']
    startDateTime = recordDateTime
    requestForConfirmation = email['RecipientsEmail']
    priority = 'Normal'
    if message.Importance == 2:
        priority = 'Higher'
    if message.Importance == 0:
        priority = 'Lower'

    ls_val = []
    ls_val.append([issueId,eventId,userName,description,recordDateTime,startDateTime,requestForConfirmation,priority,EntryID])

    ls_columns = ['IssueId','eventId','UserName','Description','RecordDateTime','StartDateTime','RequestForConfirmation','Priority','EntryID']

    df_ev = pandas.DataFrame(ls_val,columns=ls_columns, dtype='object')

    print(df_ev.head())

    df_ev.to_sql('Issue_Event', engine, if_exists='append', index=False)

    sql = "update NumberSequence set nextId = nextId + 1 where tableName = 'Issue_Event' and fieldName = 'eventId'"
    cnxn.execute(sql)
    cnxn.commit()

    return df_ev.iloc[0]

def send_email(subject, HTMLbody, recipients, attachments):

    outlookD =win32com.client.Dispatch("Outlook.Application")
    newMail = outlookD.CreateItem(0)
    newMail.Subject = subject

    recipients = sorted(recipients)

    newMail.To = ";".join(recipients)
    newMail.HTMLBody = HTMLbody

    print(recipients)

    # attach files
    for attachment in attachments:
        print(attachment)
        newMail.Attachments.Add(attachment)
        
    newMail.Send()    

    return

def reply(event):

    text_file = open("Acknowledgement.htm", "r")
    o = text_file.read()
    text_file.close()

    sql = "select * from TBL_DLOG_EMAIL where EntryID = '" + event['EntryID'] + "'"
    df = pandas.read_sql(sql,cnxn2)

    email = df.iloc[0]

    row = df.iloc[0]
    
    o = o.replace('[NAME]',email['Recipients'])
    o = o.replace('[EVENT ID]',event['eventId'])
    o = o.replace('[ISSUE ID]',event['IssueId'])
    o = o.replace('[PRIORITY]',event['Priority'].lower())       
    
    #rplyall=message.ReplyAll()
    #rplyall.HTMLBody = o +rplyall.HTMLBody 
    #rplyall.Send()

    return

outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

outlookD =win32com.client.Dispatch("Outlook.Application")

inbox=outlook.GetDefaultFolder(6)

messages=inbox.Items

#messages = inbox.Items.Restrict("[SenderEmailAddress] = '/O=EXCHANGELABS/OU=EXCHANGE ADMINISTRATIVE GROUP (FYDIBOHF23SPDLT)/CN=RECIPIENTS/CN=FA5A963B2215454B93C63ED2C4BD8504-DAVID TSANG'")

nowDate = datetime.now().strftime("%Y-%m-%d")

messages = inbox.Items.Restrict("[ReceivedTime] >= '" + nowDate + "'")

messages.Sort("[ReceivedTime]", True)

message = messages.GetFirst()

ls_message = []


while message:

    ReceivedTime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    try:
        ReceivedTime = message.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
    except:
        pass
    
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
    recipientsEmail = ''
    try:
        for recipient in message.Recipients:
            recipients = recipients + ',' + str(recipient)
            email = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
            if len(email) > 0:
                recipientsEmail = recipientsEmail  + ','  + recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
    except:
        pass

    try:
        for recipient in message.CC:
            recipients = recipients + ',' + str(recipient)
            email = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
            if len(email) > 0:
                recipientsEmail = recipientsEmail  + ','  + recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress

    except:
        pass    


    try:

        recipients = recipients[1:]
        recipientsEmail = recipientsEmail[1:]
        
        ls_message = []
        ls_message.append([message.EntryID,message.Subject,message.Body,senderEmailAddress,recipients,recipientsEmail, ReceivedTime, message.Importance])

        df_message = pandas.DataFrame(ls_message, columns=['EntryID','Subject', 'Body', 'SenderEmailAddress', 'Recipients','RecipientsEmail','ReceivedTime','Importance'])

        sql = "select * from TBL_DLOG_EMAIL where EntryID = '" + message.EntryID + "'"
        df = pandas.read_sql(sql,cnxn2)

        if len(df) == 0:
            df_message.to_sql('TBL_DLOG_EMAIL', engine2, if_exists='append', index=False)

        searchString = message.Subject
        searchString = searchString.replace("'","")

        CheckOK = True

        sql = "select * from TBL_DLOG_SenderEmailAddress where SenderEmailAddress = '" + senderEmailAddress + "'"
        df = pandas.read_sql(sql,cnxn2)
        print(df)
        if len(df) == 0:
            CheckOK = False
              
        if CheckOK == True:
            sql = "select top  1 * from [issue] where title like '%" + searchString + "%' or "
            sql = sql + "'" + searchString + "' like '%title%' order by CreateDateTime desc"
            print(sql)
            df = pandas.read_sql(sql,cnxn)

            if len(df) > 0:
                issue = df.iloc[0]
            else:
                issue = createIssue(message.EntryID)

            event = createIssueEvent(message.EntryID, issue)

            if recipientsEmail == 'david.tsang@ck-lifesciences.com':
                reply(event)

        attachments = message.attachments        

        for attachment in attachments:
            fn = attachment.FileName
            local_path = os.getcwd()
            local_path = local_path + '\\' + event['eventId']
            os.mkdir(local_path)
            local_path = os.path.join(local_path, attachment.FileName)   
            attachment.SaveAsFile(local_path)            

    except Exception as error:

        #Sending error
        subject = 'Failed to create issue for ' + message.Subject
        body = str(error) + '<br/><br/><br/>' + str(message.HTMLBody)
        recipients = ['david.tsang@ck-lifesciences.com']
        attachments = []
        send_email(subject, body, recipients, attachments)

    message = messages.GetNext()


exit()
