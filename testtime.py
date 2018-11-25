# -*-coding:utf-8-*-
from __future__ import print_function
import datetime
from timesimulation import timesimulate


def main():
    timesimulate(r'SaveWindows2018_7_31_7-57-57.db')
    time_now = datetime.datetime.now().strftime('%H:%M:%S.%f')
    time_milli = datetime.datetime.now().strftime('%f')
    print(time_now)


if __name__ == '__main__':
    main()
