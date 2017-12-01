###########
#python 2.7
###########

###
# Suunto Movement Extract
###

import os
import openpyxl
#import json

###
# Implementation Notes
# json.dumps() converts dictionary to JSON string
###

class SuuntoMovementExtract:
    debug = False

    GPSFIELDS = {
        0: "LocalTimes",
        1: "Altitude",
        2: "Distance",
        3: "SeaLevelPressure",
        4: "Speed",
        5: "Temperature",
        6: "VerticalSpeed",
        7: "Cadence",
        8: "RelativePerformanceLevel"
    }

    def __init__(self, isDebug):
        self.debug = isDebug

    def parseMoveDirectory(self):
        for sheet in os.listdir("./moves"):
            if sheet.endswith(".xlsx"):
                if self.debug: print sheet

                #Open spreadsheet
                workBook = openpyxl.load_workbook(os.path.join("./moves", sheet))
                if self.debug: print type(workBook)

                #First sheet opened by default
                workSheet = workBook.active
                if self.debug: print workSheet

                #Find "Move Samples" row and column in sheet
                cellValue, colIndex, rowIndex = "", 0, 1
                while cellValue != "Move samples":
                    colIndex += 1
                    cellValue = workSheet.cell(row=rowIndex, column=colIndex).value
                    if self.debug: print cellValue

                #Iterate thru each row. parse data into dictionary, add to sheet dictionary
                sheetGPSData = {}
                rowIndex, colOffset = 3, 0
                while cellValue != None:
                    GPSDataPoint = {}

                    #flesh out GPSDatapoint, iterate thru 9 columns of row
                    for colOffset in range(1, 9):




                #Create file, dump dictionary into sheet as JSON
                outputFilename = sheet[:-5] + "_GPSEXTRACT.txt"
                if self.debug: print outputFilename

                outputFile = open(outputFilename, "w")
                #outputFile.write(json.dumps(sheetGPSData)
                #outputFile.flush()
                #outputFile.close()



SuuntoMovementExtract(True).parseMoveDirectory()
