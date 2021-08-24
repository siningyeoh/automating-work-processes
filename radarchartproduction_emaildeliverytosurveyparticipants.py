import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import plotly.express as px
import pandas as pd

smtpObj = smtplib.SMTP('smtp-mail.outlook.com', 587)
smtpObj.ehlo()
smtpObj.starttls()
own_email=(input('What is your email address?\n'))
own_password=(input('What is your password?\n'))
smtpObj.login(own_email, own_password)

testdf=pd.read_excel('dummy dataset.xlsx')

#improve this section
df=testdf.set_index('Email').T.reset_index().rename(columns={'index':'Process'}) 
separated = df ["Process"].str.split(".", n = -1, expand = True) 
postdf = pd.concat([df, separated], axis=1).rename(columns = {1:'theta'}).drop(columns=['Process', 0]) 

for column in postdf.iloc[:,:-1]:
  testfig = px.line_polar(postdf, r=column, theta='theta', line_close=True)
  testfig.write_image("final.png")

  message = MIMEMultipart()
  message['Subject'] = "Partnership Development"
  #personalise it and send it in email body
  message.attach(MIMEText('Dear colleague, please find your survey results below.'))
  message_body = MIMEText('<br><img src="cid:%s"><br>' % ('final.png'), 'html')  
  message.attach(message_body)

  #improve this section
  file = open('final.png', 'rb')                                                    
  image = MIMEImage(file.read())
  file.close()
  image.add_header('Content-ID', '<{}>'.format('final.png'))
  message.attach(image)
  smtpObj.sendmail(own_email, column, message.as_string())

smtpObj.quit()
