#!/usr/bin/env python3
'''
Strips HTML decoration from text
'''

from bs4 import BeautifulSoup
import sys

# Example HTML content
html_content = ''
for line in sys.stdin: html_content+=line

# Use BeautifulSoup to parse the HTML
soup = BeautifulSoup(html_content, 'html.parser')

# Extract the text content from the parsed HTML
text_content = soup.get_text(separator='\n', strip=True)

print(text_content)

