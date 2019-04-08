__author__ = 'infraff0000'
import csv
from xlrd import open_workbook


# declare class
class inputOutput:
    this_file = ''
    csv_list = []

    def __init__(self):
        print "inputOutput created"

    #load file as string
    def loadFile(self, whatFile):
        with open(whatFile, "r") as myfile:
            self.this_file = myfile.read().replace('\n', '')

    def csvToList(self, whatFile):
        with open(whatFile, 'rU') as csv_file:
            reader = csv.reader(csv_file)
            for row in reader:
                self.csv_list.append(row)

    def loadXLS(self, whatFile):
        self.this_file = open_workbook(whatFile)

    def getFile(self):
        return self.this_file

    def getList(self):
        return self.csv_list



