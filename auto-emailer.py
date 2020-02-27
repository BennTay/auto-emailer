import win32com.client as win32
import mysql.connector
import json
from datetime import datetime as dt, timedelta as td
import os
import getpass

production = False

rcaInc_wrongPaste = 'Wrong label pasted'
rcaInc_wrongLogRef = 'Wrong LogRef provided'
rcaInc_wrongEan = 'Wrong EAN provided'
rcaInc_wrongScan = 'Scanned wrong code'

rcaVS_missingBoth = 'LogRef and EAN missing'
rcaVS_missingLogRef = 'LogRef not available'
rcaVS_missingEan = 'EAN not available'

class Info:
    def __init__(self, level3_correct, level3_incorrect, level3_VS, level4_correct, level4_incorrect, level4_VS, incorrectRcaList, vsRcaList):
        self.level3_correct = level3_correct
        self.level3_incorrect = level3_incorrect
        self.level3_VS = level3_VS
        self.level3_total = level3_correct + level3_incorrect + level3_VS

        self.level3_correct_percentage = '%.1f%%' % ((level3_correct/self.level3_total)*100)
        self.level3_incorrect_percentage = '%.1f%%' % ((level3_incorrect/self.level3_total)*100)
        self.level3_VS_percentage = '%.1f%%' % ((level3_VS/self.level3_total)*100)

        self.level4_correct = level4_correct
        self.level4_incorrect = level4_incorrect
        self.level4_VS = level4_VS
        self.level4_total = level4_correct + level4_incorrect + level4_VS

        self.level4_correct_percentage = '%.1f%%' % ((level4_correct/self.level4_total)*100)
        self.level4_incorrect_percentage = '%.1f%%' % ((level4_incorrect/self.level4_total)*100)
        self.level4_VS_percentage = '%.1f%%' % ((level4_VS/self.level4_total)*100)

        self.total_correct = level3_correct + level4_correct
        self.total_incorrect = level3_incorrect + level4_incorrect
        self.total_VS = level3_VS + level4_VS
        self.grandTotal = self.level3_total + self.level4_total

        self.total_correct_percentage = '%.1f%%' % ((self.total_correct/self.grandTotal)*100)
        self.total_incorrect_percentage = '%.1f%%' % ((self.total_incorrect/self.grandTotal)*100)
        self.total_VS_percentage = '%.1f%%' % ((self.total_VS/self.grandTotal)*100)

        self.incorrectRcaList = incorrectRcaList
        self.visualCheckRcaList = vsRcaList

class RCA:
    def __init__(self, status, rc, num):
        self.status = status
        self.rc = rc
        self.num = num

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
        #queryDistinct = 'select count(distinct EAN, `Putaway Label`) from %s' % (self.conn_params['validation_table'])

        clauseCorrect = ' and Status=\"CORRECT\"'
        clauseIncorrect = ' and Status=\"INCORRECT\"'
        clauseVisualCheck = ' and Status=\"VISUAL CHECK\"'
        #clauseNotValidated = ' where `Validation State`=\"Not Validated\"'

        clauseLevel3 = ' and Computer like \"%LEVEL 3%\"'
        clauseLevel4 = ' and Computer like \"%LEVEL 4%\"'

        clauseIncorrect_wrongLabelPasted = ' and `Root Cause`=\"%s\"' % (rcaInc_wrongPaste)
        clauseIncorrect_wrongLogRefProvided = ' and `Root Cause`=\"%s\"' % (rcaInc_wrongLogRef)
        clauseIncorrect_wrongEanProvided = ' and `Root Cause`=\"%s\"' % (rcaInc_wrongEan)
        clauseIncorrect_wrongScan = ' and `Root Cause`=\"%s\"' % (rcaInc_wrongScan)

        clauseVisualCheck_missBoth = ' and `Root Cause`=\"%s\"' % (rcaVS_missingBoth)
        clauseVisualCheck_missLogRef = ' and `Root Cause`=\"%s\"' % (rcaVS_missingLogRef)
        clauseVisualCheck_missEan = ' and `Root Cause`=\"%s\"' % (rcaVS_missingEan)

        ## Inspection summary
        # Level 3
        self.cur.execute(query + clauseCorrect + clauseLevel3)
        numCorrect_level3 = self.cur.fetchone()[0]

        self.cur.execute(query + clauseIncorrect + clauseLevel3)
        numIncorrect_level3 = self.cur.fetchone()[0]

        self.cur.execute(query + clauseVisualCheck + clauseLevel3)
        numVisualCheck_level3 = self.cur.fetchone()[0]

        # Level 4
        self.cur.execute(query + clauseCorrect + clauseLevel4)
        numCorrect_level4 = self.cur.fetchone()[0]

        self.cur.execute(query + clauseIncorrect + clauseLevel4)
        numIncorrect_level4 = self.cur.fetchone()[0]

        self.cur.execute(query + clauseVisualCheck + clauseLevel4)
        numVisualCheck_level4 = self.cur.fetchone()[0]

        ## INCORRECT RCA
        incorrectRcaList = []
        if numIncorrect_level3 + numIncorrect_level4 > 0:
            self.cur.execute(query + clauseIncorrect_wrongLabelPasted)
            numWrongPaste = self.cur.fetchone()[0]
            rcaObj_wrongPaste = RCA('INCORRECT', rcaInc_wrongPaste, numWrongPaste)
            incorrectRcaList.append(rcaObj_wrongPaste)

            self.cur.execute(query + clauseIncorrect_wrongLogRefProvided)
            numWrongLogRef = self.cur.fetchone()[0]
            rcaObj_wrongLogRef = RCA('INCORRECT', rcaInc_wrongLogRef, numWrongLogRef)
            incorrectRcaList.append(rcaObj_wrongLogRef)

            self.cur.execute(query + clauseIncorrect_wrongEanProvided)
            numWrongEan = self.cur.fetchone()[0]
            rcaObj_wrongEan = RCA('INCORRECT', rcaInc_wrongEan, numWrongEan)
            incorrectRcaList.append(rcaObj_wrongEan)

            self.cur.execute(query + clauseIncorrect_wrongScan)
            numWrongScan = self.cur.fetchone()[0]
            rcaObj_wrongScan = RCA('INCORRECT', rcaInc_wrongScan, numWrongScan)
            incorrectRcaList.append(rcaObj_wrongScan)

            incorrectRcaList.sort(key=lambda item: item.num, reverse=True)

        ## VISUAL CHECK RCA
        visualCheckRcaList = []
        if numVisualCheck_level3 + numVisualCheck_level4 > 0:
            self.cur.execute(query + clauseVisualCheck_missBoth)
            numMissBoth = self.cur.fetchone()[0]
            rcaObj_missBoth = RCA('VISUAL CHECK', rcaVS_missingBoth, numMissBoth)
            visualCheckRcaList.append(rcaObj_missBoth)

            self.cur.execute(query + clauseVisualCheck_missLogRef)
            numMissLogRef = self.cur.fetchone()[0]
            rcaObj_missLogRef = RCA('VISUAL CHECK', rcaVS_missingLogRef, numMissLogRef)
            visualCheckRcaList.append(rcaObj_missLogRef)

            self.cur.execute(query + clauseVisualCheck_missEan)
            numMissEan = self.cur.fetchone()[0]
            rcaObj_missEan = RCA('VISUAL CHECK', rcaVS_missingEan, numMissEan)
            visualCheckRcaList.append(rcaObj_missEan)

            visualCheckRcaList.sort(key=lambda item: item.num, reverse=True)

        return Info(level3_correct=numCorrect_level3, level3_incorrect=numIncorrect_level3, level3_VS=numVisualCheck_level3, level4_correct=numCorrect_level4, level4_incorrect=numIncorrect_level4, level4_VS=numVisualCheck_level4, incorrectRcaList=incorrectRcaList, vsRcaList=visualCheckRcaList)

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

def generate_html(infoObj, date):
    head = '<h2>Inbound report for %s</h2>' % (date)
    info = '<p>Good morning, here is the status report for yesterday\'s inbound activity:</p>'

    if infoObj.grandTotal > 0:
        tableStart = '<table border=1 frame=hsides rules=rows bordercolor=\"black\">'
        dataRow = '<tr><th align=\"center\">%s</th><td align=\"center\">%s</td><td align=\"center\">%s (%s)</td><td align=\"center\">%s (%s)</td><td align=\"center\">%s (%s)</td></tr>'
        row1 = '<tr><th>%s</th><th>%s</th><th>%s</th><th>%s</th><th>%s</th></tr>' % ('Level', 'Total no. of inspections', 'CORRECT', 'INCORRECT', 'VISUAL CHECK')
        row2 = dataRow % ('4', infoObj.level4_total, infoObj.level4_correct, infoObj.level4_correct_percentage, infoObj.level4_incorrect, infoObj.level4_incorrect_percentage, infoObj.level4_VS, infoObj.level4_VS_percentage)
        row3 = dataRow % ('3', infoObj.level3_total, infoObj.level3_correct, infoObj.level3_correct_percentage, infoObj.level3_incorrect, infoObj.level3_incorrect_percentage, infoObj.level3_VS, infoObj.level3_VS_percentage)
        row4 = dataRow % ('Total', infoObj.grandTotal, infoObj.total_correct, infoObj.total_correct_percentage, infoObj.total_incorrect, infoObj.total_incorrect_percentage, infoObj.total_VS, infoObj.total_VS_percentage)
        tableEnd = '</table>'
        additional = '<p><em>Percentages shown are in relation to respective row totals.</em></p>'
        summaryTable = additional + tableStart + row1 + row2 + row3 + row4 + tableEnd

        if infoObj.total_incorrect > 0:
            incorrectTable = '<table border=1 frame=hsides rules=rows bordercolor="black"><tr><th>%s</th><th>%s</th><th>%s</th><th>%s</th></tr><tr><td align=\"center\">%s</td><td align=\"center\">%s</td><td align=\"center\">%s</td><td align=\"center\">%s</td></tr></table>' % (infoObj.incorrectRcaList[0].rc, infoObj.incorrectRcaList[1].rc, infoObj.incorrectRcaList[2].rc, infoObj.incorrectRcaList[3].rc, infoObj.incorrectRcaList[0].num, infoObj.incorrectRcaList[1].num, infoObj.incorrectRcaList[2].num, infoObj.incorrectRcaList[3].num)
        else:
            incorrectTable = '<p><em>No INCORRECT inspections were encountered.</em></p>'
        
        if infoObj.total_VS > 0:
            visualcheckTable = '<table border=1 frame=hsides rules=rows bordercolor="black"><tr><th>%s</th><th>%s</th><th>%s</th></tr><tr><td align=\"center\">%s</td><td align=\"center\">%s</td><td align=\"center\">%s</td></tr></table>' % (infoObj.visualCheckRcaList[0].rc, infoObj.visualCheckRcaList[1].rc, infoObj.visualCheckRcaList[2].rc, infoObj.visualCheckRcaList[0].num, infoObj.visualCheckRcaList[1].num, infoObj.visualCheckRcaList[2].num)
        else:
            visualcheckTable = '<p><em>No VISUAL CHECK inspections were encountered.</em></p>'

        table = '<h3>Summary of inspections</h3>' + summaryTable + '<h3>Incorrect cases RCA</h3>' + incorrectTable + '<h3>Visual Check cases RCA</h3>' + visualcheckTable
    
    else:
        table = '<p><em>No inspections were done yesterday.</em></p>'

    misc = '<p>If there are any issues or enquiries, please contact me through <a href=\"mailto:bennguobin.tay@se.com\">e-mail</a> or <a href=\"lync15:bennguobin.tay@se.com?chat\">Skype</a> and the human me will respond. Thank you!</p>'
    signoff = '<h4>Sincerely,<br>Benn\'s Inbound Report Bot</h4>'
    htmlStr = head + info + table + misc + signoff
    return htmlStr

def main():
    outlook = win32.Dispatch('outlook.application')

    try:
        user = getpass.getuser()
        directory = os.path.join('C:\\', 'Users', user, 'Box', 'Inbound Auto Report')
        if production:
            path_connParams = os.path.join(directory, 'conn_params_viralcom.json')
            path_mailList = os.path.join(directory, 'mail_list.json')
        else:
            path_connParams = 'conn_params_viralcom.json'
            path_mailList = 'mail_list.json'

        with open(path_connParams, 'r') as f:
            conn_params = json.load(f)
        connector = DBConnector(conn_params)
        
        mailList = ''
        with open(path_mailList, 'r') as f:
            mailDict = json.load(f)
            for recipientName in mailDict:
                mailList += mailDict[recipientName]
        
        ytd = dt.now() - td(days=1)
        ytdStr = dt.strftime(ytd, '%Y-%m-%d')
        ytdNiceDate = dt.strftime(ytd, '%d %b %Y')
        infoObj = connector.retrieve_data(ytdStr)
        htmlStr = generate_html(infoObj, ytdNiceDate)
        #totalNumScans, numCorrect, numIncorrect, numVisualCheck, incorrectRcaList, visualCheckRcaList, numNotValidated = connector.retrieve_data(ytdStr)
        #htmlStr = generate_html(totalNumScans, numCorrect, numIncorrect, numVisualCheck, ytdNiceDate, incorrectRcaList, visualCheckRcaList, numNotValidated)

        mail = outlook.CreateItem(0)
        mail.To = mailList
        mail.CC = 'bennguobin.tay@se.com'
        mail.Subject = '%s Inbound Inspection Report' % (ytdNiceDate)
        mail.HTMLBody = htmlStr
        mail.Send()
    
    except Exception as e:
        print(str(e))
        send_error_email(outlook, str(e))

    finally:
        connector.disconnect()
        print('DBConnector disconnected.')


main()
