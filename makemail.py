#!/usr/bin/env python3

'''
Construct email task from list of people

Arguments: (may be abbreviated as first letter)
   -list XMLFILE : list of names and other attributes as needed
   -directory XMLFILE : directory of email addresses
   -attribute NAME VALUE : to choose subset of names based on attributes
   -parameter NAME VALUE : to set additional attributes as needed by email template
   -outfile XMLFILE : otherwise writes to stdout
'''

import sys
import xml.etree.ElementTree as ET

def read_xml(filename):
   tree = ET.parse(filename)
   return tree.getroot()

def parse_xml(filename):
   data=[]
   for child in read_xml(filename): data+=[child.attrib] # converts to dict
   return data

args=iter(sys.argv[1:])
addr=[]; attr={}; param={}; dir=[]; outfile=''
for arg in args:
   if arg=='-list' or arg=='-l': addr+=parse_xml(next(args))
   elif arg=='-directory' or arg=='-d': dir+=parse_xml(next(args))
   elif arg=='-attribute' or arg=='-a':
      ai=next(args); at=next(args)
      attr[ai]=at # if use next(args) directly, gives wrong order
   elif arg=='-parameter' or arg=='-p':
      pi=next(args); pt=next(args)
      param[pi]=pt
   elif arg=='-outfile' or arg=='-o': outfile=next(args)

dir1={}
for di in dir: dir1[di['name']]=di

root = ET.Element("mail"); root.tail='\n'
tree = ET.ElementTree(root)
for addri in addr:
   process=True
   for attri in attr.keys():
      if addri[attri]!=attr[attri]: process=False
   if process:
      newel = ET.Element("rcpt"); newel.tail='\n'
      newel.attrib.update(addri)
      newel.attrib.update(param)
      name=addri['name']
      newel.attrib['email']=dir1[name]['email']
      names=name.split()
      newel.attrib['firstname']=names[0]
      newel.attrib['lastname']=names[-1]
      root.append(newel) 

if outfile=='': ET.dump(tree)
else: tree.write(outfile)
