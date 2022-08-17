from flask import Flask, render_template, request
from werkzeug.utils import secure_filename
from werkzeug.datastructures import  FileStorage
import pandas as pd
#import  jpype     
import  asposecells   
import vobject
import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import datetime
#jpype.startJVM()




ALLOWED_EXTENSIONS = {'xls','xlsx'}

app = Flask(__name__)

def write_vcard(f, vcard):
    with open(f, 'w') as f:
        f.writelines([l + '\n' for l in vcard])


def converti(file_xls): 
    
   # from asposecells.api import Workbook
    workbook = Workbook(file_xls)
    workbook.save("file_converted.xlsx")
    
    


def invia_email(numerocontatti):
  
  date = datetime.date.today()
  datet = date.strftime("%d %B %Y")
  subject = "Contatti " + str(datet)
  body="Ecco i tuoi " + str(numerocontatti) + " Contatti"
  sender_email = "noleggio.bot@gmail.com"
  receiver_email = "pasqualepepe887@gmail.com"
  password = "kfzlshkkgzadiahf"
  message = MIMEMultipart()
  message["From"] = sender_email
  message["To"] = receiver_email
  message["Subject"] = subject
  message["Bcc"] = receiver_email 
  message.attach(MIMEText(body, "plain"))

  filename = "contacts.vcf"  # In same directory as script
  with open(filename, "rb") as attachment:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())

# Encode file in ASCII characters to send by email    
  encoders.encode_base64(part)

# Add header as key/value pair to attachment part
  part.add_header(
    "Content-Disposition",
    f"attachment; filename= {filename}",
  )

# Add attachment to message and convert message to string
  message.attach(part)
  text = message.as_string()

# Log in to server using secure context and send email
  context = ssl.create_default_context()
  with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(sender_email, password)
    server.sendmail(sender_email, receiver_email, text)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/', methods = ['GET', 'POST'])
def hello():
   

   if request.method == 'POST':
      f = request.files['file']
      if f and allowed_file(f.filename):

         f.save(secure_filename(f.filename))
         #converti(f.filename)
         df = pd.read_excel(f.filename, usecols="o,p,n,m")   #o = full_name / p = phone
         rows_count = df.count()[0]
         print(df)
         vcard = [""]
    # print(rows_count)
    ### print(df.iat[0,1])
         for i in range(rows_count):

             phone = df.iat[i,3]
             name = df.iat[i,2]
             km=df.iat[i,1]
             auto = df.iat[i,0]
      # print(name, phone)
             phone = "+" + str(phone)
             vcard = vcard +['BEGIN:VCARD',
             'VERSION:2.1',
             f'N:{name}',
             f'TEL;WORK;VOICE:{phone}',
             f'TITLE:{km}',
             f'ORG:{auto}',
             f'REV:1',
             'END:VCARD']







   
         write_vcard('contacts.vcf', vcard)
         invia_email(rows_count)
         return "<style> body{background-color: lightblue;}</style><h1>Contatti inviati correttamente...</h1>"
     

   
   return render_template('index.html')
   
	
if __name__ == '__main__':
   app.run(debug = False,host="0.0.0.0")
  # jpype.shutdownJVM()	

   
