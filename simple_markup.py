#!/usr/bin/python
#coding=utf-8

import sys,re

from util import *

print '<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><title>...</title></head><body>'

title = True
for block in blocks(sys.stdin):
    block = re.sub(r'\*(.+?)\*',r'<u><font color="red">\1</font></u>',block)
    if title:
        print '<h1>'
        print block
        print '</h1>'

        title = False
    else:
        print '<p>'
        print block
        print '</p>'

print '</body></html>'
