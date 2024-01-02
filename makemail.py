#!/usr/bin/env python3

import sys
import xml.etree.ElementTree as ET

def read_xml(filename):
   tree = ET.parse(filename)
   return tree.getroot()

def parse_xml(filename):
   data=[]
   for child in read_xml(filename): data+=[child.attrib] # converts to dict
   return data

def save_xml_file(tree, filename): tree.write(filename)

args=iter(sys.argv[1:])
addr=[]; type={}; param={}; dir=[]; outfile=''
for arg in args:
   if arg=='-addressees' or arg=='-a': addr+=parse_xml(next(args))
   elif arg=='-directory' or arg=='-d': dir+=parse_xml(next(args))
   elif arg=='-type' or arg=='-t':
      ti=next(args); tt=next(args)
      type[ti]=tt # if use next(args) directly, gives wrong order
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
   for typei in type.keys():
      if addri[typei]!=type[typei]: process=False
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
