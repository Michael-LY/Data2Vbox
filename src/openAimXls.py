# -*- encoding: utf-8 -*-
#!/usr/bin/python3


import csv
import os
import sys
import openpyxl

def readfile(filename):
    '''Print a file to the standard output.'''
    f = open(filename)
    while True:
        line = f.readline()
        if len(line) == 0:
            break
        print(line)
    f.close()

if __name__ == '__main__':
    readfile(sys.argv[1])

