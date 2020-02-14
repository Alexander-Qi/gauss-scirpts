# import time
import datetime
# import re
import os
import sys
# import math
# import numpy
import xlwt
debug = "no"

scriptversion = 0.32
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
# hartree to kcal/mol and kJ/mol
har2kcal = 627.5095294
har2kj = 2625.499871

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
            if file.endswith(".log") and not("irc" in file):
                filelist.append(file)
    # If using full or relative path for file or dir
    elif "/" in sys.argv[1]:
        # Checking if a single file with path
        if ".log" in sys.argv[1]:
            dirmode = "off"
            filename = sys.argv[1]
            filelist.append(filename)
            print("filename is", filename)
        # Or a directory
        else:
            dirmode = "on"
            for file in os.listdir(sys.argv[1]):
                if file.endswith(".log"):
                    filelist.append(sys.argv[1]+"/"+file)
    # If parent folder
    elif sys.argv[1] == "..":
        dirmode = "on"
        for file in sorted(os.listdir(sys.argv[1])):
            if file.endswith(".log"):
                filelist.append(sys.argv[1]+"/"+file)
    else:
        dirmode = "off"
        filename = sys.argv[1]
        if '.log' in filename == None:
            print("This script can only read \".log\" Gaussian output files, you can try altering the filename extension.")
            exit()
        filelist.append(filename)
except IndexError:
    print(bcolors.OKBLUE + "Gaussian Output reader version",
          scriptversion, "(Python version)", bcolors.ENDC)
    print("---------------------------------")
    print("Script usage:")
    print("On single file: \n G-output-reader.py orcafile.out\n or python -u G-output-reader.py orcafile.out")
    print("On directory: \n G-output-reader .\n or python -u G-output-reader.py .")
    quit()


book = xlwt.Workbook()
sheet = book.add_sheet(sheetname='sheet 1')
output_path = 'Results-' + today + '.xls'

sheet.write(0, 0, "Task Name")
sheet.write(0, 1, "Zero-point correction")
sheet.write(0, 2, "Thermal correction to Energy")
sheet.write(0, 3, "Thermal correction to Enthalpy")
sheet.write(0, 4, "Thermal correction to Gibbs Free Energy")
sheet.write(0, 5, "Sum of electronic and zero-point Energies")
sheet.write(0, 6, "Sum of electronic and thermal Energies")
sheet.write(0, 7, "Sum of electronic and thermal Enthalpies")
sheet.write(0, 8, "Sum of electronic and thermal Free Energies")
sheet.write(0, 9, "Opt time")
# sheet.write(0, 11, "Hartree to kcal/mol")
# sheet.write(0, 12, har2kcal)
# sheet.write(1, 11, "Hartree to kJ/mol")
# sheet.write(1, 12, har2kj)


linenum = 0
for filename in filelist:
    linenum = linenum + 1
    zpcorr = "not found"
    TcorrEne = "not found"
    TcorrEnth = "not found"
    TcorrGibbs = "not found"
    Szp = "not found"
    Sene = "not found"
    Senth = "not found"
    Sgibbs = "not found"
    taskname = filename.split(".")[0]
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
            if 'Zero-point correction' in line:
                zpcorr = float(line.split()[2])
            if 'Thermal correction to Energy' in line:
                TcorrEne = float(line.split()[4])
            if 'Thermal correction to Enthalpy' in line:
                TcorrEnth = float(line.split()[4])
            if 'Thermal correction to Gibbs Free Energy' in line:
                TcorrGibbs = float(line.split()[6])
            if 'Sum of electronic and zero-point Energies' in line:
                Szp = float(line.split()[6])
            if 'Sum of electronic and thermal Energies' in line:
                Sene = float(line.split()[6])
            if 'Sum of electronic and thermal Enthalpies' in line:
                Senth = float(line.split()[6])
            if 'Sum of electronic and thermal Free Energies' in line:
                Sgibbs = float(line.split()[7])
            if 'Job cpu time:' in line:
                coreclock_day = float(line.split()[3])
                coreclock_hour = float(line.split()[5])
                coreclock_minute = float(line.split()[7])
                coreclock_second = float(line.split()[9])
                coreclock = float(((coreclock_day * 24) + coreclock_hour +
                                  (coreclock_minute / 60) + (coreclock_second / 3600)))

    # writing into an .xls file
    sheet.write(linenum, 0, taskname)
    sheet.write(linenum, 1, zpcorr)
    sheet.write(linenum, 2, TcorrEne)
    sheet.write(linenum, 3, TcorrEnth)
    sheet.write(linenum, 4, TcorrGibbs)
    sheet.write(linenum, 5, Szp)
    sheet.write(linenum, 6, Sene)
    sheet.write(linenum, 7, Senth)
    sheet.write(linenum, 8, Sgibbs)
    sheet.write(linenum, 9, coreclock)


book.save(output_path)
