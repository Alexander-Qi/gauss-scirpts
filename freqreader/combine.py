# Calculation Results Combiner
import openpyxl
import datetime
import os
import sys
import pandas
# import xlrd
# import xlwt
from xlrd import open_workbook
from xlwt import Workbook
from xlutils.copy import copy

debug = "no"

har2kcal = 627.5095294
har2kj = 2625.499871

scriptversion = 0.3

today = str(datetime.date.today())


minimaloutput = 0
if sys.argv[1] == "short":
    minimaloutput = 1
elif sys.argv[1] == "long":
    minimaloutput = 0


freqdata = pandas.DataFrame(pandas.read_excel(
    "E:\\opt\\" + 'Results-' + today + '.xls', sheet_name='sheet 1'))

zpandasata = pandas.DataFrame(pandas.read_excel(
    "E:\\zeropoint\\"+'Results-zp-' + today + '.xls', sheet_name='sheet 1'))

# freqdata = xlrd.open_workbook("/opt/"+'Results-' + today + '.xls').sheets()[0]

# zpandasata = xlrd.open_workbook("/zeropoint/"+'Results-zp-' + today + '.xls').sheets()[0]
# result = pandas.merge(freqdata, zpandasata.loc[:, ['Task Name', 'Zero-Point Time']], how='right', on='Task Name')
result = pandas.merge(freqdata, zpandasata.loc[:, [
    'Task Name', 'Zero-Point Time', 'Final Zero-point Energy']], how='right', on='Task Name')

# result.eval("Final Gibbs Free Energy = Thermal correction to Gibbs Free Energy + Final Zero-point Energy" , inplace=True)
result["Total Time"] = result[["Opt time", "Zero-Point Time"]
                              ].apply(lambda x: x["Opt time"] + x["Zero-Point Time"], axis=1)
result["Final Gibbs Free Energy(kcal)"] = result[["Thermal correction to Gibbs Free Energy", "Final Zero-point Energy"]
                                           ].apply(lambda x: 627.5095294*(x["Thermal correction to Gibbs Free Energy"] + x["Final Zero-point Energy"]), axis=1)

if minimaloutput == 0:
    intermed = pandas.ExcelWriter('All-Results-' + today + '.xls')
    result.to_excel(intermed, index=False)
    intermed.save()

    raw = open_workbook('All-Results-' + today + ".xls")
    out = copy(raw)

    outsheet = out.get_sheet(0)

    outsheet.write(0, 15, "Hartree to kcal/mol")
    outsheet.write(0, 16, har2kcal)
    outsheet.write(1, 15, "Hartree to kJ/mol")
    outsheet.write(1, 16, har2kj)

    outsheet.col(4).width = len("Thermal correction to Gibbs Free Energy")*128
    # outsheet.col(10).width=len("Final Zero-point Energy")*256
    outsheet.col(13).width = len("Final Gibbs Free Energy")*256
    outsheet.col(15).width = len("Hartree to kcal/mol")*256

    out.save('All-Results-' + today + '.xls')
elif minimaloutput == 1:
    miniresult = result.drop(['Zero-point correction', 'Thermal correction to Energy', 'Thermal correction to Enthalpy', 'Sum of electronic and zero-point Energies',
                              'Sum of electronic and thermal Energies', 'Sum of electronic and thermal Enthalpies', 'Sum of electronic and thermal Free Energies', 'Opt time', 'Zero-Point Time'], axis=1, inplace=True)
    intermed2 = pandas.ExcelWriter('Short-Results-' + today + '.xls')
    result.to_excel(intermed2, index=False)
    intermed2.save()

    raw2 = open_workbook('Short-Results-' + today + ".xls")
    out2 = copy(raw2)

    out2sheet = out2.get_sheet(0)

    out2sheet.write(0, 6, "Hartree to kcal/mol")
    out2sheet.write(0, 7, har2kcal)
    out2sheet.write(1, 6, "Hartree to kJ/mol")
    out2sheet.write(1, 7, har2kj)

    out2sheet.col(1).width = len("Thermal correction to Gibbs Free Energy")*128
    # outsheet.col(10).width=len("Final Zero-point Energy")*256
    out2sheet.col(4).width = len("Final Gibbs Free Energy")*256
    out2sheet.col(6).width = len("Hartree to kcal/mol")*256

    out2.save('Short-Results-' + today + '.xls')
