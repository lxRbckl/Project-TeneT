# Project TeneT by Alex Arbuckle #


# Import <
from json import load
from xlutils.copy import copy
from xlrd import open_workbook
from xlwt import Workbook, easyxf
#from os import startfile as start

# >


# Declaration <
with open('TeneT.json', 'r') as fileVariable:

    setting = load(fileVariable)

# >


def createFile():
    '''  '''

    # Create <
    wb = Workbook()
    ws = wb.add_sheet('worksheet')

    # >

    # Write <
    [ws.write((c + 2), 3, i, easyxf(setting['wheelSpeedStyle'])) for c, i in enumerate(setting['wheelSpeedValue'])]
    ws.write(1, 4, 'Offset', easyxf(setting['offsetStyle']))
    ws.write(1, 5, 'Delta', easyxf(setting['deltaStyle']))
    ws.write(1, 1, 'Input', easyxf(setting['inputStyle']))

    # >

    # Save <
    wb.save('TeneT.xls')

    # >


def readFile():
    '''  '''

    # Open <
    wb = open_workbook('TeneT.xls')
    ws = wb.sheet_by_index(0)

    # >

    # Read <
    listVariable = []
    tupleVariable = (1, 4, 5)
    for i in tupleVariable:

        j, var = 0, []
        while (True):

            try:

                value = ws.cell_value((j + 2), i)
                [var.append(int(value)) if (value != '') else (None)]
                j += 1

            except:

                listVariable.append(var)
                break

    # >

    # Return <
    return listVariable

    # >


def writeFile(arg):
    ''' list(list()) '''

    # Create File <
    wb = copy(open_workbook('TeneT.xls', formatting_info = True))
    ws = wb.get_sheet(0)

    # >

    # iterate 2D <
    for r, i in enumerate(arg):

        # iterate 1D <
        for c, j in enumerate(i):

            ws.write((r + 2), (c + 6), j)

        # >

    # >

    # Save <
    wb.save('TeneT.xls')

    # >


# Main <
if (__name__ == '__main__'):

    # Create File <
    input('< Hit Enter to Create File >')

    createFile()
    #start('TeneT.xls')

    # >

    # Read File <
    input('< Hit Enter to Write File >')

    listVariable = []
    inputList, offsetList, deltaList = readFile()
    multiTable = setting['multiplicationTableValue']
    for i in range(len(offsetList)):

        var = []
        for j in range(len(multiTable)):

            formula = str(((offsetList[i] * (int(multiTable[j]) * int(deltaList[j]))) % 38))
            [var.append(int(formula)) for i in range(inputList[j])]

        listVariable.append(var)

    # >

    # Write File <
    writeFile(listVariable)
    #start('TeneT.xls')

    # >
