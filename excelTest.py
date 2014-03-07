import xlrd
import re

book = xlrd.open_workbook("testDataSheet2.xls")

backgroundSheetName = "Background"

backgroundSheet = book.sheet_by_name(backgroundSheetName)


headerRow = 3

#print taskSheet.ncols


participants = []


def getBackgroundSheet():
    pass

def getParticipants():
    global participants
    for rowNumber in range(2, backgroundSheet.nrows):
        testDate = backgroundSheet.cell_value(rowNumber, 2)
        if testDate:
            newParticipant = {}
            for colNumber, colValue in enumerate(backgroundSheet.row_values(1)):
                #print colValue
                if colValue != "":
                    if backgroundSheet.cell_type(rowNumber, colNumber) == 1: # text
                        newParticipant[colValue] = backgroundSheet.cell_value(rowNumber, colNumber)
                    elif backgroundSheet.cell_type(rowNumber, colNumber) == 2: # float
                        newParticipant[colValue] = backgroundSheet.cell_value(rowNumber, colNumber)
                    elif backgroundSheet.cell_type(rowNumber, colNumber) == 3: # date
                        newParticipant[colValue] = backgroundSheet.cell_value(rowNumber, colNumber)
                    else:
                        newParticipant[colValue] = backgroundSheet.cell_value(rowNumber, colNumber)
                    newParticipant[colValue] = backgroundSheet.cell_value(rowNumber, colNumber)
                    print "newParticipant[",colValue,"]: ", newParticipant[colValue]
                participants.append(newParticipant)
    return participants        
    

def getUseErrors():
    useErrors = []
    for sheet in book.sheets():
        for rownum in range(1, sheet.nrows):
            pt = []
            row_values = sheet.row_values(rownum)
            
            for colNumber, value in enumerate(row_values):
                try:
                    searchResult = re.search('\[[uU][eE]\][^[]*', value)
                    if searchResult is not None:
                        if sheet.cell(0, 0).value[0] == "T" and sheet.cell(0, 1).value is not "P#":
                            newUseError = {}
                            newUseError["task"] = sheet.cell(0, 0).value
                            newUseError["participant"] = sheet.cell(rownum, 0).value
                            newUseError["description"] = searchResult.group(0)
                            newUseError["useErrorName"] = sheet.cell(1, colNumber).value
                            useErrors.append(newUseError)
                            print "---"
                            print "Task: ", newUseError["task"]
                            print "Participant: ", newUseError["participant"]
                            print "Use Error: ", newUseError["useErrorName"]
                            print "Description: ", newUseError["description"]
                except Exception as e:
                    pass
                    #print "ERROR: ", str(e)
                    #print ""   

getUseErrors()

getParticipants()

# regex for finding use errors: \[[uU][eE]\][^[]*
'''
for rownum in range(1, taskSheet.nrows):
    pt = []
    row_values = taskSheet.row_values(rownum)
    
    for value in row_values:
        try:
            searchResult = re.search('\[[uU][eE]\][^[]*', value)
            if searchResult is not None:
                print "---"
                print "Task: ", taskSheet.cell(0, 0).value
                print "Participant: ", taskSheet.cell(rownum, 0).value
                print searchResult.group(0)
        except Exception:
            doNothing = ""
            #print ""
        
        '''
    
    
    
    #for colNum in range(1, taskSheet.ncols):
    #   print taskSheet.row_values(headerRow)[colNum]
        
    
    
    


