#!/usr/bin/env python3

'''
Manage form emails by controlling gmail via SMTP and IMAP.
Access via app password, not OAuth API.
- Sends emails to a list of addressees, filling out a template message
  with addressee-specific material.
  Addressees are defined in an XML file.
  Addressee-specific parameters are attributes under each entry.
  Attribute names are referred to in the template as an uppercase keyword with a leading $,
  e.g. $FIRSTNAME for attribute "firstname".
  Sent messages are logged to another XML file.
- Checks for replies, logging the reply date and moving the reply to a chosen folder.
- Send a reminder for any message with no reply logged, using a reminder template.

Multiple files may be specified for the addressees and the email template.
The lists and template contents are concatenated.

The email template must have:
- A line starting with 'To: ' followed by the attribute name for the email address.
- A line starting with 'Subject: ' followed by the subject, which can contain ad
dressee-specific keywords.
- 'Body:' on a line by itself, followed by the body of the email (containing keywords) on subsequent lines.

Additional text files can be inserted in the body of the email template.
If a line starts with '++', the rest of the line is interpreted as a filename
e.g. "++signature.txt".
Substitution is recursive: inserted files are processed for '++' commands too.
The script checks for insertion *after* replacing keyword fields,
so any filenames in the template may themselves be set to addressee-specific values by using keywords.

Arguments:
   -template/-t TEMPLATE_FILE (needed for send and remind only)
   -addressees/-a XML_FILE (needed for send only)
   -folder/-f MAIL_FOLDER (default: automail)
   -log/-l LOGFILE (default: log.xml)
   -username/-u GMAIL_USER
   -account GMAIL_ACCOUNT
   -password/-p GMAIL_PASSWORD
   -passvar GMAIL_PASSWORD_ENVIRONMENT_VARIABLE
   -send/-s
   -check/-c (checks for replies)
   -recheck (checks even if a reply has already been received)
   -remind/-r
   -debug/-d (Prints send or remind messages but doesn't send or log. Specify *before* -send/-remind argument.)
'''

import sys,re,subprocess,time,datetime,copy
import os,smtplib
import imaplib
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import decode_header
import xml.etree.ElementTree as ET

def expandfile(file):
   text=''
   with open(file,'r') as fp:
      for line in fp:
         if line.startswith('++'): text+=expandfile(line[2:])
         else: text+=line
   return text

def subwords(text,vars):
   text1=text
   for var in vars.keys():
      var1=r'\$'+var.upper() # keyword is $ plus varname in uppercase
      val=vars[var]
      text1=re.sub(var1,val,text1)
   text2=''
   for line in text1.splitlines():
      if line.startswith('++'):text2+=expandfile(line[2:])
      else: text2+=line+'\n'
   return text2

def readfile(file):
   with open(file,'r') as fp: return fp.read()

def read_xml(filename):
   tree = ET.parse(filename)
   return tree.getroot()

def parse_xml(filename):
   data=[]
   for child in read_xml(filename): data+=[child.attrib] # converts to dict
   return data

def logtime(tt=None):
   if tt==None: tt=datetime.datetime.now()
   return tt.strftime("%-m/%-d/%y %H:%M")

def readlogdata(logfile):
   # read log file
   try:
      tree = ET.parse(logfile)
      root = tree.getroot()
   except FileNotFoundError:
      print(f"File '{logfile}' not found, will create.")
      root = ET.Element("root"); root.tail='\n'
      tree = ET.ElementTree(root)
   except ET.ParseError: raise(f"Error: File '{logfile}' could not be parsed.")
   except Exception as e: raise(f"Unexpected error occurred: {e}")
   return(tree,root)

def logmsg(logfile,msg,vars={}):
   (tree,root)=readlogdata(logfile)

   # create new item
   newel = ET.Element("msg"); newel.tail='\n'
   newel.attrib.update(vars)
   addr=getaddress(msg)
   subj=getsubject(msg)
   newel.set('address',addr)
   newel.set('subject',subj)
   newel.set('sent',logtime())
   newel.set('reply','')

   # append and write
   root.append(newel)
   tree.write(logfile)   

def getline(text,label):
   for line in text.splitlines():
      if line.startswith(label): return line[len(label):]
   return ''

def getaddress(text): return getline(text,'To: ')
def getsubject(text): return getline(text,'Subject: ')

def getbody(text):
   body=''
   bodymode=False
   for line in text.splitlines():
      if bodymode: body+=line+'\n'
      elif line.startswith('Body:'): bodymode=True
   return body

def send_gmail(msg):
   global user,acct,password
   addr=getaddress(msg)
   subj=getsubject(msg)
   body=getbody(msg)

   fromaddr=acct+'@gmail.com'
   gmsg = MIMEMultipart()
   gmsg['From'] = f'{user} <{fromaddr}>'
   gmsg['To'] = addr
   gmsg['Subject'] = subj
   gmsg.attach(MIMEText(body, 'plain'))

   try:
      server = smtplib.SMTP('smtp.gmail.com', 587)
      server.starttls()
      server.login(acct,password)
      server.sendmail(fromaddr, addr, gmsg.as_string())
      server.quit()
      return True
   except Exception as e:
      print("Error: ", e)
      return False

def send(template,addressees,log):
   global debug
   for addressee in addressees:
      msg=subwords(template,addressee)
      if debug: print(msg)
      elif send_gmail(msg): logmsg(log,msg,addressee)

def gmail_getreply(addr,subj,savefolder):
   global acct,password

   imap = imaplib.IMAP4_SSL("imap.gmail.com")
   imap.login(acct, password)

   # create savefolder if necessary
   status, folders = imap.list()
   if status == 'OK':
      folder_exists = any(f'"{savefolder}"' in folder.decode() for folder in folders)
   
      if not folder_exists:
         # Folder doesn't exist, so create it
         print(f"Creating folder: {savefolder}")
         imap.create(savefolder)

   imap.select("inbox")
   status, messages = imap.search(None, f'(FROM "{addr}" SUBJECT "{subj}")')
   messages = messages[0].split() # convert result to list of email IDs

   reply=''
   for mail in messages:
      _, msg = imap.fetch(mail, "(RFC822)")
      for response in msg:
         if isinstance(response, tuple):
            msg = email.message_from_bytes(response[1]) # parse raw content
            subject = decode_header(msg["Subject"])[0][0]
            if isinstance(subject, bytes): subject = subject.decode()
            reply+=f'From: {addr}\nSubject: {subj}\n'

            # extract and print body
            if msg.is_multipart():
               for part in msg.walk():
                  if part.get_content_type() == "text/plain":
                     body = part.get_payload(decode=True).decode()
                     reply+=f'Body:\n{body}\n'
            else:
               body = msg.get_payload(decode=True).decode()
               reply+=f'Body:\n{body}\n'

            # move to save folder (label)
            imap.store(mail, '+X-GM-LABELS', savefolder)
            imap.store(mail, '+FLAGS', '\\Deleted')
            imap.expunge()

   imap.close()
   imap.logout()
   print(reply)
   return reply

def check(log,savefolder,all=False):
   (tree,root)=readlogdata(log)
   change=False
   for msg in root:
      msgreply=msg.get('reply')
      if all or msgreply=='':
         addr=msg.attrib['address']
         subj=msg.attrib['subject']
         reply=gmail_getreply(addr,subj,savefolder)
         if reply!='':
            print('Reply:'); print(reply)
            if msgreply!='': msgreply+=','
            msgreply+=logtime()
            msg.set('reply',msgreply)
            change=True
   if change: tree.write(log)

def remind(template,log):
   global debug
   (tree,root)=readlogdata(log)
   change=False
   for logitem in root:
      if logitem.attrib['reply']=='':
         vars=copy.deepcopy(logitem.attrib)
         vars['subject']='Re: '+vars['subject']
         msg=subwords(template,vars)
         if debug: print(msg)
         elif send_gmail(msg):
            remind=''
            if 'remind' in logitem.attrib: remind=logitem.attrib['remind']+','
            remind+=logtime()
            logitem.set('remind',remind)
            change=True
   if change: tree.write(log)

if __name__=='__main__':
   template=''
   addressees=[]
   log='log.xml'
   mailfolder='automail'
   user=acct=passvar=password=''
   debug=False
   args=iter(sys.argv[1:])
   for arg in args:
      if arg=='-template' or arg=='-t': template+=readfile(next(args)) # appends
      elif arg=='-addressees' or arg=='-a': addressees+=parse_xml(next(args))
      elif arg=='-log' or arg=='-l': log=next(args)
      elif arg=='-username' or arg=='-u': user=next(args)
      elif arg=='-account': acct=next(args)
      elif arg=='-password' or arg=='-p': password=next(args)
      elif arg=='-passvar': password=os.environ.get(next(args))
      elif arg=='-send' or arg=='-s': send(template,addressees,log)
      elif arg=='-folder' or arg=='-f': mailfolder=next(args)
      elif arg=='-check' or arg=='-c': check(log,mailfolder)
      elif arg=='-recheck': check(log,mailfolder,all=True)
      elif arg=='-remind' or arg=='-r': remind(template,log)
      elif arg=='-debug' or arg=='-d': debug=True
      else: raise ValueError(f'Argument {arg} not recognized!')
