import os

from flask import Flask, render_template, request, flash
import pandas as pd
from werkzeug.utils import secure_filename, redirect

app = Flask(__name__)

UPLOAD_FOLDER = '/c/Users/PRABHAT/PycharmProjects/flaskApi/uploads/'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


#---------Date&Time-Block-----------
from datetime import datetime
now = datetime.now()
time_date=now.strftime("%d%m")
print(time_date)
##-----------------------------------


ALLOWED_EXTENSIONS = {'xlsx'}
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def index():
    return render_template('index.html')

@app.route('/data', methods=['GET', 'POST'])
def data():
    if request.method == 'POST':
        if request.files:

            file = request.files['upload-file']
            ToEmail=request.form['email']
            print(ToEmail)

        # if file.filename == '':
        #     flash('No selected file')
        #     return redirect(request.url)
        # if file and allowed_file(file.filename):
        #     filename = secure_filename(file.filename)

            APP_ROOT = os.path.dirname(os.path.abspath(__file__))

            # Current directory for Flask app + file name
            # Use this file_path variable in your code to refer to your file
            file_path = os.path.join(APP_ROOT, file.filename)

            file.save(os.path.join(APP_ROOT, file.filename))
            print("Saved")
            print(file_path)

            print(os.path.splitext(file.filename)[0])
            # return redirect(request.url)
            import numpy as np
            import pandas as pd
            import requests
            import json

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
            Config for postgres
    
    
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            def getanswer(req):
                headers = {'content-type': 'application/json'}
                payload = req
                res = requests.post('https://abcl.vitt.ai',
                                    data=json.dumps(payload), headers=headers)
                if res:
                    res = res.json()
                else:
                    res = None
                return res


            exceldf = pd.read_excel(
                open(file_path, 'rb'), sheet_name='Sheet1')
            writedf = pd.DataFrame(columns=['id', 'question', 'dbanswer',
                                            'apianswer', 'apiintent', 'apientity', 'apientityvalue', 'match'])
            cnt = 0
            req = {"session": -1, "query": "hi!",
                   "type": "sent", "time": "xyz", "count": 1, "conversationId": "yolo"}

            for i, row in exceldf.iterrows():
                # if str(row['answer']).lower() == 'nan':
                #     continue
                req['query'] = row['question']
                res = getanswer(req)
                if not res:
                    continue
                listofmessages = res.get('result').get('fulfillment').get('messages')
                if '' in listofmessages:
                    listofmessages.remove('')
                apianswer = []

                try:
                    entitylist = [key for key, value in res.get(
                        'result').get('parameters').items()]
                except:
                    entitylist = []
                try:
                    entityvaluelist = [value for key, value in res.get(
                        'result').get('parameters').items()]
                    entityvaluelist = [
                        item for sublist in entityvaluelist for item in sublist]
                except:
                    entityvaluelist = []
                try:
                    capturedintent = res.get('result').get('metadata').get('intentName')
                except:
                    capturedintent = None

                for item in listofmessages:
                    if item.get('type') == 0:
                        if item.get('speech')[0] != '':
                            apianswer.append(item.get('speech')[0])

                writedf.at[cnt, 'id'] = cnt+1
                writedf.at[cnt, 'question'] = row['question']
                writedf.at[cnt, 'dbanswer'] = row['answer']
                writedf.at[cnt, 'apianswer'] = '|'.join(apianswer)
                writedf.at[cnt, 'apiintent'] = capturedintent
                if entitylist:
                    writedf.at[cnt, 'apientity'] = ','.join(entitylist)
                if entityvaluelist:
                    writedf.at[cnt, 'apientityvalue'] = ','.join(entityvaluelist)

                if str(row['answer']).lower() == 'nan':
                    dbmessages = ['k']
                else:  dbmessages = row['answer'].split('|')
                # print("137")
                # print(dbmessages)

                while '' in dbmessages:
                    dbmessages.remove('')

                if not apianswer:
                    writedf.at[cnt, 'match'] = 'nomatch'
                    cnt += 1
                    continue
                if apianswer[0][:10] == dbmessages[0][:10]:
                    writedf.at[cnt, 'match'] = 'match'
                else:
                    writedf.at[cnt, 'match'] = 'nomatch'

                cnt += 1
                if cnt % 10 == 0:
                    print(cnt)

            writer = pd.ExcelWriter(os.path.splitext(file.filename)[0]+"_output_"+time_date+'.xlsx')
            writedf.to_excel(writer, 'Sheet1', index=False)
            writer.save()

            ##########################################################ExcelChecker41
            import smtplib
            from email.mime.multipart import MIMEMultipart
            from email.mime.text import MIMEText
            from email.mime.base import MIMEBase
            from email import encoders

            fromaddr = 'excelchecker41@gmail.com'
            toaddr = ToEmail

            # instance of MIMEMultipart
            msg = MIMEMultipart()

            # storing the senders email address
            msg['From'] = fromaddr

            # storing the receivers email address
            msg['To'] = toaddr

            # storing the subject
            msg['Subject'] = "Excel file"

            # string to store the body of the mail
            body = "Check attachment"

            # attach the body with the msg instance
            msg.attach(MIMEText(body, 'plain'))

            # open the file to be sent
            filen =  os.path.splitext(file.filename)[0]+"_output_"+time_date+'.xlsx'
            attachment = open(filen,  "rb")

            # instance of MIMEBase and named as p
            p = MIMEBase('application', 'octet-stream')

            # To change the payload into encoded form
            p.set_payload((attachment).read())

            # encode into base64
            encoders.encode_base64(p)

            p.add_header('Content-Disposition', "attachment; filename= %s" % filen)

            # attach the instance 'p' to instance 'msg'
            msg.attach(p)

            # creates SMTP session
            s = smtplib.SMTP('smtp.gmail.com', 587)
            #
            # start TLS for security
            s.starttls()



            # Converts the Multipart msg into a string
            text = msg.as_string()

            # Authentication
            s.login(fromaddr, "checkerexcel41")

            # sending the mail
            s.sendmail(fromaddr, toaddr, text)



            return '''
                successfully Done!check your Email
                '''

if __name__ == '__main__':
    app.run(debug=True)
