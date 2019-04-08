__author__ = 'infraff0000'
from xlrd import open_workbook
from xlwt import Workbook
from tempfile import TemporaryFile


class XLSProcessor:
    orig_source = ''
    source_list = []
    source_dictionary = {}

    def __init__(self):
        print "XLSProcessor created"

    def setOrigSource(self, passed_source):
        print "setOrigSource"
        self.orig_source = passed_source

    def hasOrigSource(self):
        if not self.orig_source:
            print " orig_source Not Set"
            return 0
        else:
            print " orig_source Set"
            return 1

    def hasSourceList(self):
        if not self.source_list:
            print " List Not Set"
            return 0
        else:
            print " source_list Set"
            return 1

    def loadXLS(self, whatFile):
        print "loadXLS file: " + whatFile
        self.orig_source = open_workbook(whatFile)

    def compileMasterSourceIntoDictionary(self):
        print "compileAgileSourceIntoDictionary"
        # take the workbook and turn it into a dictionary

        if self.hasOrigSource():
            # get a copy of the sheet, indexed
            sheet = self.orig_source.sheet_by_index(0)

            for row_index in range(1, sheet.nrows):
                # print the result of this row, for feedback
                # print sheet.row_values(row_index)
                self.source_list.append(sheet.row_values(row_index))

            self.source_dictionary = dict((x[0], x) for x in self.source_list)
            print "Master Dictionary created length of: "
            print len(self.source_dictionary)

    def returnDictionaryItemBySKU(self, SKU):
        if SKU in self.source_dictionary:
            return self.source_dictionary[SKU]
        else:
            return False

    def removeDuplicatesFromList(self):
        print "removeDuplicatesFromList"
        if self.hasSourceList():
            #self.orig_source = list(set(self.orig_source))
            # remove duplicates based on index 0
            self.source_dictionary = dict((x[1], x) for x in self.source_list)
            print "Source Dictionary created length of: "
            print len(self.source_dictionary)

    def isolateTargetedRows(self, start):
        # in the excel file, the rows start at 18 (or 17, because row 1  = 0 )

        # get a copy of the sheet, indexed
        sheet = self.orig_source.sheet_by_index(0)

        # iterate from the starting row
        for row_end in range(start, sheet.nrows):
            # get row number at the end of the list
            if type(sheet.cell(row_end, 1).value) is not unicode:
                # print the number of the ending row
                print "row " + str(row_end) + " should be the last row"
                # now that we got it, stop
                break

        # now, for only the range we are interested in
        for row_index in range(start, row_end):
            # print the result of this row, for feedback
            #print sheet.row_values(row_index)

            # add the current row to the source_list list
            self.source_list.append(sheet.row_values(row_index))

    def indexSourceBySku(self):
        #save as index by SKU
        for l in self.orig_source:
            self.source_dictionary[l[0]] = l

    def printXLS(self):
        for s in self.orig_source.sheets():
            print 'Sheet:', s.name
            for row in range(s.nrows):
                values = []
                for col in range(s.ncols):
                    values.append(s.cell(row, col).value)
                print ','.join(values)

    def clearItems(self):
        print "clearItems"
        self.orig_source = ''
        self.source_dictionary = {}
        self.source_list = []

    def saveResultlistAsXLS(self, result_list, path):
        print "saveResultlistAsXLS"
        book = Workbook()
        sheet1 = book.add_sheet('Sheet 1')

        row_number = 0
        for row in result_list:
            sheet1.write(row_number, 0, row[0])
            sheet1.write(row_number, 1, row[1])
            sheet1.write(row_number, 2, row[2])
            sheet1.write(row_number, 3, row[3])
            sheet1.write(row_number, 4, row[4])

            row_number += 1

        book.save(path)
        book.save(TemporaryFile())
        print "booked saved as " + path
