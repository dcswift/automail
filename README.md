# Email automation tools
e.g. for coordinating related messages to/from multiple people

Versions for gmail and MS Outlook on Mac, 
and a tool to create a list of emails to coordinate.

## Usage
Lists are supplied as XML files, with data read from element attributes.

The email-coordinating scripts take a list of addressees and a template email,
and use attributes from the list to set addressee-specific fields in the template.
They create a log file (also XML) of messages sent, and use it to check for replies and record reminders sent.

It's easy to accidentally send a lot of erroneous emails using automated tools,
so these scripts have a debug mode that should be used to check operation
before sending emails for real.
The debug mode just prints out the messages that would be sent.

In real situations, email exchanges to be coordinated involve different subsets of a group of people and different information.
The list coordinating tool <code>makemail</code> constructs the addressee list with appropriate attributes 
from a separate list of people in a group and a directory list of email addresses (also XML).
People can be included or excluded by different group attributes (e.g. "member" vs "associate")
and additional attributes can be specified for the particular set of messages sent.

Example:
<pre>
makemail.py -l team.xml -d directory.xml -a status member -p project project\ name -o addressees.xml
outlookmac.py -a addressees.xml -t template1.txt [-debug] -send
outlookmac.py -check
outlookmac.py -t template1r.txt [-debug] -remind
</pre>

## Installation
To install: git clone --depth 1 https://github.com/dcswift/automail.git
