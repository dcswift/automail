#!/usr/bin/env python3

'''
Manage form emails by controlling MS Outlook using AppleScript on Mac OSX.
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
- A line starting with 'Subject: ' followed by the subject, which can contain addressee-specific keywords.
- 'Body:' on a line by itself, followed by the body of the email (containing keywords) on subsequent lines.

Additional text files can be inserted in the body of the email template.
If a line starts with '++', the rest of the line is interpreted as a filename
e.g. "++signature.txt". 
Substitution is recursive: inserted files are processed for '++' commands too.
The script checks for insertion *after* replacing keyword fields,
so any filenames in the template may themselves be set to addressee-specific values by using keywords.

Arguments: (may be abbreviated as first letter)
   -template TEMPLATE_FILE (needed for send and remind only)
   -html (format email body as HTML, which seems necessary to preserve line breaks with Outlook)
   -addressees XML_FILE (needed for send only)
   -folder MAIL_FOLDER (default: automail)
   -log LOGFILE (default: log.xml)
   -send
   -check
   -remind
   -debug (Prints send or remind messages but doesn't send. Specify *before* -send/-remind argument.)
'''

import sys,re,subprocess,time,datetime,copy
import xml.etree.ElementTree as ET

def subwords0(text, vars): # replace text words if in dict
   words = text.split()
   words1 = [vars[word] if word in vars else word for word in words]
   return ' '.join(words1)

def subwords1(text,vars): # replace text words if in dict
   lines = text.splitlines()
   text1 = ''
   for line in lines:
      words = line.split()
      words1 = []
      for word in words:
         if word.isupper(): words1+=[vars.get(word.lower(),word)]
         else: words1+=[word]
      line1 = ' '.join(words1)
      text1+=line1+'\n'
   return text1

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

def applescript_email(msg):
   global html
   addr=getaddress(msg)
   subj=getsubject(msg)
   body=getbody(msg)
   if html: body='<body>'+body.replace('\n','<br>\n')+'</body>'
   script = f'''
tell application "Microsoft Outlook"
    set newMessage to make new outgoing message with properties {{subject:"{subj}", content:"{body}"}}
    tell newMessage
        make new recipient at newMessage with properties {{email address:{{address:"{addr}"}}}}
        send
    end tell
end tell
'''
   return script

def send(template,addressees,log):
   global debug
   for addressee in addressees:
      msg=subwords(template,addressee)
      if debug: print('Message:\n',msg)
      run_applescript_str(applescript_email(msg))
      logmsg(log,msg,addressee)

def applescript_getreply(addr,subj,savefolder):
    applescript = f'''
tell application "Microsoft Outlook"
    set output to ""
    set targetFolder to missing value

    -- Check if the folder exists, if not, create it
    repeat with aFolder in mail folders
        if name of aFolder is "{savefolder}" then
            set targetFolder to aFolder
            exit repeat
        end if
    end repeat
    if targetFolder is missing value then
        set targetFolder to make new mail folder at end of mail folders with properties {{name: "{savefolder}"}}
    end if

    -- Search for matching messages
    set theMessages to messages of inbox
    repeat with aMessage in theMessages
        try
            set senderObj to sender of aMessage
            set senderAddress to address of senderObj as string
            set senderName to name of senderObj
            set subjectLine to subject of aMessage
            if senderAddress is "{addr}" and subjectLine contains "{subj}" then
                set output to "Subject: " & subjectLine & return & "From: " & senderName & return & "Content: " & content of aMessage
                move aMessage to targetFolder
                exit repeat
            end if
            on error errMsg number errNum
                set output to output & "Error processing a message: " & errMsg & ", Error Number: " & errNum & return
        end try
    end repeat
    return output
end tell
'''
    return applescript

def check(log,savefolder):
   (tree,root)=readlogdata(log)
   change=False
   for msg in root:
      if msg.attrib['reply']=='':
         addr=msg.attrib['address']
         subj=msg.attrib['subject']
         reply=run_applescript_str(applescript_getreply(addr,subj,savefolder))
         reply1=''
         for line in reply.splitlines():
            if line.startswith('Error processing'): print(line)
            else: reply1+=line+'\n'
         if reply1!='':
            print('Reply:'); print(reply) # complete with error messages
            msg.set('reply',logtime())
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
         if debug: print('Message:\n',msg)
         run_applescript_str(applescript_email(msg))
         remind=''
         if 'remind' in logitem.attrib: remind=logitem.attrib['remind']+','
         remind+=logtime()
         logitem.set('remind',remind)
         change=True
   if change: tree.write(log)

def run_applescript_file(script_path):
   try: subprocess.run(['osascript', script_path], check=True)
   except subprocess.CalledProcessError as e: print(f"AppleScript Error: {e}")

def run_applescript_str(script):
   global debug
   if debug:
      print('script:\n',script)
      return ''
   process = subprocess.Popen(['osascript', '-'], stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
   stdout, stderr = process.communicate(script)
   #print('osascript output:\n',stdout)
   if process.returncode == 0: return stdout.strip()
   else: raise Exception(f"AppleScript Error: {stderr}")

if __name__=='__main__':
   template=''
   addressees=[]
   log='log.xml'
   mailfolder='automail'
   html=debug=False
   args=iter(sys.argv[1:])
   for arg in args:
      if arg=='-template' or arg=='-t': template+=readfile(next(args)) # appends
      elif arg=='-html' or arg=='-h': html=True
      elif arg=='-addressees' or arg=='-a': addressees+=parse_xml(next(args))
      elif arg=='-log' or arg=='-l': log=next(args)
      elif arg=='-send' or arg=='-s': send(template,addressees,log)
      elif arg=='-folder' or arg=='-f': mailfolder=next(args)
      elif arg=='-check' or arg=='-c': check(log,mailfolder)
      elif arg=='-remind' or arg=='-r': remind(template,log)
      elif arg=='-script': run_applescript_file(next(args))
      elif arg=='-debug' or arg=='-d': debug=True
      else: raise ValueError(f'Argument {arg} not recognized!')
