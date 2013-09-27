#!/bin/python

from HTMLParser import HTMLParser
from re import sub



"""
'id', 'results.column-heading-part_data_18')
('id', 'results.column-heading-manuf_data_18')
('id', 'results.column-heading-description_data_18')
('id', 'results.column-heading-avail_data_18')
('id', 'results.column-heading-list-price_data_
"""

result = []

class LinksParser(HTMLParser):
    def __init__(self):
        HTMLParser.__init__(self)
        self.st = False
        self.data = {}
        self.key = None
        self.price_td = False

    def handle_starttag(self, tag, attr):
        if self.price_td:
            self.key = 'price'
            self.st = True

        if tag == 'table' or self.key is not None:
            for i in attr:
                if i[0] == 'class' and i[1] == 'PriceListContent':
                    self.price_td = True
                    self.st = True
                    self.key = 'price'

        if tag == 'td' or self.key is not None:
            for i in attr:
                if i[0] == 'id':
                    if "results.column-heading-part_data" in i[1] or self.key == 'part':
                        self.st = True
                        self.key = 'part'
                    elif "results.column-heading-manuf_data" in i[1] or self.key == 'manuf':
                        self.st = True
                        self.key = 'manuf'
                    elif "results.column-heading-description_data" in i[1] or self.key == 'desc':
                        self.st = True
                        self.key = 'desc'
                    elif "results.column-heading-avail_data" in i[1] or self.key == 'avail':
                        self.st = True
                        self.key = 'avail'
                    elif "results.column-heading-list-price_data" in i[1] or self.key == 'price':
                        self.key = 'price'
                        self.st = True

    def handle_endtag(self, tag):
        if tag == 'table':
            self.price_td = False

        if tag == 'td':
            self.st = False

        if not self.price_td and not self.st:
            if self.data:
                result.append(self.data)
                self.data = {}
                self.key = None

            self.last = False


    def handle_data(self, data):
        if self.st:
            if data.strip():
                """
                if self.key == 'price':
                    print self.key,
                    print data.strip()
                    print
                """

                if self.key is not None:
                    if self.data.has_key(self.key):
                        self.data[self.key] += " " + str(data.strip())
                    else:
                        self.data[self.key] = data.strip()



HELP_MSG = """Usage %s type value package"""

import urllib2
import sys
import res_table
import cap_table

URL = "http://it.farnell.com/jsp/search/browse.jsp?N={PLACEHOLDER}&Ns=P_PRICE_FARNELL_IT%7C0&locale=it_IT_ENGLISH"


if __name__ == "__main__":
    limit = 4
    value = ''
    package = ''
    precision = ''
    attr = []

    if 'help' in sys.argv or '-h' in sys.argv:
        print HELP_MSG % (sys.argv[0])
        sys.exit(1)

    if len(sys.argv) < 2 or (not (sys.argv[1] in ['res','cap'])):
        print HELP_MSG % (sys.argv[0])
        sys.exit(1)

    for a in sys.argv[1:]:
        if 'limit' in a:
            a = a.strip()
            limit = int(a.split('=')[1])
            print "limit:", limit
        else:
            attr.append(str(a.strip()))


    cmd = attr[0] if len(attr) > 1 else ''
    value = attr[1] if len(attr) > 2 else ''
    package = attr[2] if len(attr) > 3 else ''

    query_str = "718+502"

    if cmd == 'res':
        query_str = "+215515"
        precision = attr[3] if len(attr) > 4 else ''

        for i in res_table.RES_TABLE:
            if i[1] == value:
                print "value:", i
                query_str += "+"+i[0]+"+"
        for i in res_table.PACKAGE_TABLE:
            if i[1] == package:
                print "package:", i
                query_str += "+"+i[0]+"+"
        for i in res_table.PRECISION_TABLE:
            if i[1] == precision:
                print "precision:", i
                query_str += "+"+i[0]+"+"

    if cmd == 'cap':
        query_str += '+215745'
        volt = attr[3] if len(attr) > 4 else ''
        diel = attr[4] if len(attr) >= 5 else ''

        for i in cap_table.VALUE_TABLE:
            if i[1] == value:
                print "value:", i
                query_str += "+"+i[0]+"+"
        for i in cap_table.PACKAGE_TABLE:
            if i[1] == package:
                print "package:", i
                query_str += "+"+i[0]+"+"
        for i in cap_table.VOLTAGE_TABLE:
            if i[1] == volt:
                print "volt:", i
                query_str += "+"+i[0]+"+"

        for i in cap_table.DIELETRIC_TABLE:
            if i[1] == diel:
                print "dieletric:", i
                query_str += "+"+i[0]+"+"

    URL = URL.replace("{PLACEHOLDER}", query_str)
    print URL

    data_in = urllib2.urlopen(URL)
    parser = LinksParser()
    parser.feed(data_in.read())

    row = 0
    print "-" * 80
    for k in result[:limit * 5]:
        for j in k.keys():
            if j == 'price':
                s = k[j]
                i = s.find('\n')
                k[j] = s[i+1:]

            if j == 'avail':
                s = k[j]
                i = s.find('T')
                k[j] = s[:i-1]

            if j == 'desc':
                s = k[j]
                i = s.find('Ult')
                k[j] = s[:i-1]

            print "%8s: %s" % (j, k[j])

        row += 1
        if row == 5:
            print "-" * 80
            row = 0

