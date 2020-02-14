# import time
import datetime
# import re
import os
import sys
# import math
# import numpy
import xlwt
debug = "no"

scriptversion = 0.17
# Designed for Gaussian opt-freq jobtypes, can read the Thermal corrections
# Please make sure the jobs are normally terminated


# start_time = time.time()
today = str(datetime.date.today())


# Reverse read function from orcajobcheck.py
# Default buffersize was 4096. 20480 works better
def reverse_lines(filename, BUFSIZE=20480):
    # f = open(filename, "r")
    filename.seek(0, 2)
    p = filename.tell()
    remainder = ""
    while True:
        sz = min(BUFSIZE, p)
        p -= sz
        filename.seek(p)
        buf = filename.read(sz) + remainder
        if '\n' not in buf:
            remainder = buf
        else:
            i = buf.index('\n')
            for L in buf[i+1:].split("\n")[::-1]:
                yield L
            remainder = buf[:i]
        if p == 0:
            break
    yield remainder


# Conversion factors
# hartree to kcal/mol
harkcal = 627.50946900


class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


##########################################
# Getting user arguments first
#########################################
# Read in filename or dir as argument
filelist = []
try:
    if sys.argv[1] == ".":
        dirmode = "on"
        for file in sorted(os.listdir(sys.argv[1])):
            if file.endswith(".out") and not(file.endswith("smd.out")):
                filelist.append(file)
    # If using full or relative path for file or dir
    elif "/" in sys.argv[1]:
        # Checking if a single file with path
        if ".out" in sys.argv[1]:
            dirmode = "off"
            filename = sys.argv[1]
            filelist.append(filename)
            print("filename is", filename)
        # Or a directory
        else:
            dirmode = "on"
            for file in os.listdir(sys.argv[1]):
                if file.endswith(".out"):
                    filelist.append(sys.argv[1]+"/"+file)
    # If parent folder
    elif sys.argv[1] == "..":
        dirmode = "on"
        for file in sorted(os.listdir(sys.argv[1])):
            if file.endswith(".out"):
                filelist.append(sys.argv[1]+"/"+file)
    else:
        dirmode = "off"
        filename = sys.argv[1]
        if '.out' in filename == None:
            print("This script can only read \".out\" ORCA output files, you can try altering the filename extension.")
            exit()
        filelist.append(filename)
except IndexError:
    print(bcolors.OKBLUE + "Gaussian Output reader version",
          scriptversion, "(Python version)", bcolors.ENDC)
    print("---------------------------------")
    print("Script usage:")
    print("On single file: \n orca-zpreader.py orcafile.out\n or python -u orca-zpreader.py orcafile.out")
    print("On directory: \n orca-zpreader .\n or python -u orca-zpreader.py .")
    quit()


book = xlwt.Workbook()
sheet = book.add_sheet(sheetname='sheet 1')
output_path = 'Results-zp-' + today + '.xls'

sheet.write(0, 0, "Task Name")
sheet.write(0, 1, "Final Zero-point Energy")
sheet.write(0, 2, "Zero-Point Time")



linenum = 0
for filename in filelist:
    linenum = linenum + 1
    finalzp = "not found"
    taskname = filename.split("_")[0]
    coreclock = "error"
    coreclock_day = 0
    coreclock_hour = 0
    coreclock_minute = 0
    coreclock_second = 0

    # reading part
    rcount = 0
    with open(filename, errors='ignore') as file:
        for line in reverse_lines(file):
            rcount = rcount+1
            if 'FINAL SINGLE POINT ENERGY' in line:
                finalzp = float(line.split()[4])
            if 'TOTAL RUN TIME:' in line:
                coreclock_day = float(line.split()[3])
                coreclock_hour = float(line.split()[5])
                coreclock_minute = float(line.split()[7])
                coreclock_second = float(line.split()[9])
                coreclock = float(16*((coreclock_day * 24) + coreclock_hour +
                                  (coreclock_minute / 60) + (coreclock_second / 3600)))

    # writing into an .xls file
    sheet.write(linenum, 0, taskname)
    sheet.write(linenum, 1, finalzp)
    sheet.write(linenum, 2, coreclock)


book.save(output_path)
