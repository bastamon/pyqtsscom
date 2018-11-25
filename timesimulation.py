# -*-coding:utf-8-*-
from __future__ import print_function
import sys
import re
import sqlite3
import openpyxl
import time
import datetime


def timesimulate(filename, millisec=datetime.datetime.now().strftime('%f')):
    temp = ''.join(re.findall(r'[^A-Za-z]', filename, flags=re.IGNORECASE))
    timestr = temp + millisec
    date = datetime.datetime.strptime(timestr, '%Y_%m_%d_%H-%M-%S.%f')
    print(date)
    return date


def mian():
    sys.argv.append('')

    if len(sys.argv) == 2:
        timesimulate(sys.argv[1])
    elif len(sys.argv) >= 3:
        timesimulate(sys.argv[1], sys.argv[2])


if __name__ == '__main__':
    mian()
