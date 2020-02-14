import win32com.client as win32
import mysql.connector
import json
from datetime import datetime as dt, timedelta as td

class DBConnector:
    def __init__(self, conn_params):
        self.conn_params = conn_params
        self.conn = None
        self.cur = None
        self.connected = False
        self.tries = 0
        

        self.conn = mysql.connector.connect(user=self.conn_params['target_user'],
                                            password=self.conn_params['target_password'],
                                            host=self.conn_params['target_host'],
                                            database=self.conn_params['target_db'])
        self.cur = self.conn.cursor()
        self.connected = True
        print('connection to [', self.conn_params['target_host'], '] successful')

    def retrieve_data(self, date):
        query = 'select count(*) from %s where cast(Datetime as Date)=\"%s\"' % (self.conn_params['records_table'], date)
        clauseCorrect = ' and Status=\"CORRECT\"'
        clauseIncorrect = ' and Status=\"INCORRECT\"'
        clauseVisualCheck = ' and Status=\"VISUAL CHECK\"'

        self.cur.execute(query)
        totalNumScans = self.cur.fetchone()[0]

        self.cur.execute(query + clauseCorrect)
        numCorrect = self.cur.fetchone()[0]

        self.cur.execute(query + clauseIncorrect)
        numIncorrect = self.cur.fetchone()[0]

        self.cur.execute(query + clauseVisualCheck)
        numVisualCheck = self.cur.fetchone()[0]

        return (totalNumScans, numCorrect, numIncorrect, numVisualCheck)

    def disconnect(self):
        self.conn.close()
        self.conn = None
        self.cur = None
        self.connected = False

# In the case of any error, sends an e-mail informing of the error.
def send_error_email(outlook, errorMsg):
    mail = outlook.CreateItem(0)
    mail.To = 'bennguobin.tay@se.com;'
    mail.Subject = 'Auto-email bot Inbound report error'
    mail.Body = 'Error message:\n' + errorMsg
    mail.Send()

def generate_html(totalNumScans, numCorrect, numIncorrect, numVisualCheck, date):
    head = '<h2>Inbound report for %s</h2>' % (date)
    info = '<p>Good morning, these are the status reports for yesterday:</p>'
    table = '<table border=1 frame=hsides rules=rows><tr><th>%s</th><th>%s</th><th>%s</th><th>%s</th></tr><tr><th>%s</th><th>%s</th><th>%s</th><th>%s</th></tr></table>' % ('Total number of inspections', 'CORRECT inspections', 'INCORRECT inspections', 'VISUAL CHECKS required', totalNumScans, numCorrect, numIncorrect, numVisualCheck)
    signoff = '<p>Sincerely, Benn\'s Inbound report bot</p>'
    htmlStr = head + info + table + signoff
    return htmlStr

def main():
    outlook = win32.Dispatch('outlook.application')

    try:
        with open('conn_params_viralcom.json', 'r') as f:
            conn_params = json.load(f)
        connector = DBConnector(conn_params)
        #mailList = 'bennguobin.tay@se.com;'
        
        mailList = ''
        with open('mail_list.json', 'r') as f:
            mailDict = json.load(f)
            for recipientName in mailDict:
                mailList += mailDict[recipientName]
        
        ytd = dt.now() - td(days=1)
        ytdStr = dt.strftime(ytd, '%Y-%m-%d')
        ytdNiceDate = dt.strftime(ytd, '%d %b %Y')
        totalNumScans, numCorrect, numIncorrect, numVisualCheck = connector.retrieve_data(ytdStr)
        htmlStr = generate_html(totalNumScans, numCorrect, numIncorrect, numVisualCheck, ytdNiceDate)

        mail = outlook.CreateItem(0)
        mail.To = mailList
        mail.Subject = 'TEST: %s Inbound Inspection Report' % (ytdNiceDate)
        mail.HTMLBody = htmlStr
        mail.Send()
    
    except Exception as e:
        print(str(e))
        send_error_email(outlook, str(e))

    finally:
        connector.disconnect()
        print('DBConnector disconnected.')


main()

