#for speak:python text-to-speech
import pyttsx3
# for read the file
import pandas as pd
#for mailing purpose
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
#tkinter:to open GUI
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import getpass
#for keep password safe: ******
from prompt_toolkit import prompt
import time



pyttsx3.speak("welcome to my program")
print("\t      --------------------------" , end="") 
print("\n\t\tWELCOME TO MY PROGRAM" )
print("\t      --------------------------") 


pyttsx3.speak("Please Provide me credentials")
email_user =  input("Enter your email : ")
#take password but not visible
email_password = prompt("Password: ", is_password=True)
pyttsx3.speak("Let me authenticate")
print("\n")

Tk().withdraw()
while True:
    
    what = input("What type of file you want to send\n(text or attachment or both) :")
    print("\n")
    msg = MIMEMultipart()
    #tkinter open and take attachment
    if what=="attachment":
     filename=askopenfilename()
     attachment =open(filename,'rb')
     part = MIMEBase('application','octet-stream')
     part.set_payload((attachment).read())
     encoders.encode_base64(part)
     part.add_header('Content-Disposition',"attachment; filename= "+filename)
     msg.attach(part)  
    
    #only send text
    elif what=="text":
     filename='text'
     attachment=input("Type your text : ")
     msg.attach(MIMEText(attachment, 'plain'))

    #both: text and attachment
    elif what=="both":
     filename='text'
     attachment=input("Type your text : ")
     msg.attach(MIMEText(attachment, 'plain'))
     filename=askopenfilename()
     attachment =open(filename,'rb')
     part = MIMEBase('application','octet-stream')
     part.set_payload((attachment).read())
     encoders.encode_base64(part)
     part.add_header('Content-Disposition',"attachment; filename= "+filename)
     msg.attach(part)  
     
    else:
      pyttsx3.speak("ohh , invalid try again")
      print('ohh , invalid try again')
      time.sleep(3)
      quit()
      

    #requirement bulk mailing or single...
    send_type=input("give me requirement bulk mailing or single mail \n type bulk or single:")
    print("\n")
    if("single"in send_type):
      type=input("send type - 'To' 'Bcc' 'Cc' :")
      email_send = input("Enter the receptant email : ")
      print("\n")
      subject = input("Give me Subject of mail : ")
      msg['From'] = email_user
      msg[type] = ", ".join(email_send) 
      msg['Subject'] = subject
      
    else:
      type=input("send type - 'To' 'Bcc' 'Cc' :")
      pyttsx3.speak("ok ready for bulk mailing")
      print("ok ready for bulk mailing..\n")
      print("select ur file name in .csv:")
      opentk=askopenfilename()
      time.sleep(3)
      e = pd.read_excel(opentk)
      col=input("enter coloumn name in file which have list of mails:")
      emails = e[col].values
      #print(emails)
      server = smtplib.SMTP('smtp.gmail.com',587)
      server.starttls()
      server.login(email_user,email_password)

      subject = input("Give me Subject of mail : ")
      msg['From'] = email_user
      msg[type] = ", ".join(emails) 
      msg['Subject'] = subject
  
  
      text = msg.as_string()
      #by loop we send mail to all
      for email in emails:
        server.sendmail(email_user,email,text)

      pyttsx3.speak("congratulation your mail has been sent successfully ")
      print("       ----------------" ,end="") 
      print("\n\t email sent!!" )
      print("       ----------------") 
      quit() 

    #mail sending
    text = msg.as_string()
    #port no : 587
    server = smtplib.SMTP('smtp.gmail.com',587)
    server.starttls()
    server.login(email_user,email_password)
    server.sendmail(email_user,email_send,text)

    pyttsx3.speak("congratulation your mail has been sent successfully ")
    print("       ----------------" ,end="") 
    print("\n\t email sent!!" )
    print("       ----------------") 
    time.sleep(3)
 
    pyttsx3.speak("what next ....continue or you want to exit ")
    p = input("continue or exit :")
    print("\n")
    if("continue" in p):
      print("       ------" ,end="") 
      print("\n\t OK" )
      print("       ------") 
      pyttsx3.speak("ok")
      continue
    else:
      print("       --------" ,end="") 
      print("\n\t BYE" )
      print("       --------") 
      pyttsx3.speak("bye")
      server.quit()
      break

        
       
  
