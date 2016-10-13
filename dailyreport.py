import  MySQLdb
import xlwt
import smtplib
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import time

def get_data(sql,**dbinfo):
    conn = MySQLdb.connect(**dbinfo)
    cur = conn.cursor()
    cur.execute(sql)
    result = cur.fetchall()
    cur.close()
    conn.close
    return result
def write_date_to_excel(name,sql,**dbinfo):
    result = get_data(sql,**dbinfo)
    sheet = wbk.add_sheet(name,cell_overwrite_ok=True)
    for i in xrange(len(result)):
       	for j in xrange(len(result[i])):
            sheet.write(i,j,result[i][j])
def post_mail(**mail):
    msg = MIMEMultipart()
    msg["Subject"] = mail["mailtitle"]
    msg["From"] = mail["sender"]
    msg["To"] = mail["reciver"]
    content = MIMEText(mail["text"])
    msg.attach(content)
    for a in mail["attachments"]:
        part = MIMEApplication(open(a,'rb').read())
        part.add_header("Content-Disposition","attachment",filename = a)
        msg.attach(part)
    server = smtplib.SMTP(mail["smtp"],mail["port"],timeout=30)
    try:
        server.login(mail["user"],mail["passwd"])
        server.sendmail(mail["sender"],mail["reciver"],msg.as_string())
    except:
       print("post fail!!")          
    finally:
        server.close()
if __name__ == '__main__':
    dbinfo1 = {"host":"",
              "user":"",
              "passwd":"",
              "db":"",
              "port":3306,
            "charset":"utf8",}
    while True:
        start_time = datetime.datetime.now()
        today = datetime.date.today()
        yesterday = today - datetime.timedelta(days=1)
        now = list(time.localtime())[3:6]
        if now[0] == 8 and now[1] == 0 and now[2] == 0:
            title_sql = {"new_register":"SELECT count(*) from sys_user u WHERE u.create_time >= '%s' and u.create_time < '%s';" % (yesterday,today),
                        }
            mail1 = {"sender":"",
                    "reciver":"",
                    "attachments":[".xls",],
                    "smtp":"",
                    "port":25,
                    "user":"",
                    "passwd":"",
                    "text":"daily report",
                    "mailtitle":"daily report(%s)" % yesterday,
                    }
            wbk = xlwt.Workbook()
            for k,v in title_sql.items():
                write_date_to_excel(k,v,**dbinfo1)
            wbk.save("dailyreport"+".xls")
            post_mail(**mail1)
            end_time = datetime.datetime.now()
            diff = end_time - start_time
            sleep_time = abs(86400-diff.total_seconds())
            time.sleep(sleep_time)
        else:
            continue
#    wbk = xlwt.Workbook()
#    for k,v in title_sql.items():
#        write_date_to_excel(k,v,**dbinfo1)
#        wbk.save("dailyreport"+".xls")
#    post_mail(**mail1)