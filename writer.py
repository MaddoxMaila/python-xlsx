import xlsxwriter as xwriter

'''
    Opened WorkBook, instance of xwriter.WorkBook
'''
workBook = xwriter.Workbook('./new files/output.xlsx')

'''
    Hold All WorkSheets For This WorkBook
'''
workSheetsList = []

'''
    Currently Opened WorkSheet
'''
currentWorkSheet = None #workBook.add_worksheet()

'''
    Holder To Hold All Specially Formatted Data Collected From Input File
'''
formattedData = []

'''
    Format Styling For The Merging
'''
mergeFormat = workBook.add_format({
    'bold' : 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': 'white'})

'''
    Get WorkBook To Use
'''
def getWorkBook(filename) :

    '''
    :param filename: Path To File
    :return: xwriter.WorkBook Instance
    '''
    return workBook


'''
    Add WorkSheet To The Currently Opened WorkBook
'''
def addWorkSheet(sheet) :

    '''
    :param sheet: Worksheets For The WorkBook
    :return: None

        This Is The Brain Of The Program. Most Stuff Happens Here!

        1. Firstly Add A WorkSheet To The WorkBook
        2. Add Headers To Specify Which Type Of Data Is
        3. Add The Data!

    '''

    '''
        Add WorkSheet To Work
    '''
    currentWorkSheet = workBook.add_worksheet(sheet)

    '''
        Dictionary To Hold The WorkSheet Name & Instance As Key/Value.
        During Iteration, We'll Just Use The Sheet Name ** :param sheet ** To Choose Which
        WorkSheet To Use
    '''
    # SheetDict = {
    #     sheet   : currentWorkSheet
    # }

    '''
        Add All WorkSheet To A Global Array
    '''
    workSheetsList.append(currentWorkSheet)

    '''
        Add Headers For Dynamic Data!
    '''

    dateColumnIndex = 1
    paymentColumn   = 2

    currentWorkSheet.write(0, 0, 'CashFlowSchedule', mergeFormat)
    currentWorkSheet.write(0, dateColumnIndex, 'Date', mergeFormat)
    currentWorkSheet.write(0, paymentColumn, 'Payment', mergeFormat)

    currentWorkSheet.merge_range("A1:A{}".format(len(formattedData) + 1), "CashFlowSchedule")

    '''
        Iterate Through The Formatted Data To Retrieve The Rows Stored In 
        This Format {"sheet_name" : "xxx", "date" : "xxx", "data" : "xxx"}
    '''

    rowIndex = 0

    for singleRowData in formattedData :

        '''
            Check If Sheet Name Passed As Argument Is Same As Sheet Name Stored
        '''
        if sheet == singleRowData['sheet_name'] :

            ''' True :=> Continue With Operation '''

            '''
                Add Date
            '''
            currentWorkSheet.write(rowIndex, dateColumnIndex, singleRowData['date'])

            '''
                Add Payment
            '''
            currentWorkSheet.write(rowIndex, paymentColumn, singleRowData['payment'])

        ''' Increment The Row Index! '''
        rowIndex += 1


'''
    Close Currently Opened WorkBook.
    
    **** Any Changes Made To The File Won't Save Without Calling This Method
'''
def workBookClose() :

    '''
    :return: None
    '''

    '''
        Close The WorkBook
    '''
    workBook.close()

'''
    Write To The The Current WorkSheet
'''
def writeToWorkingSheet(colomn, row, data) :

    """
    :param colomn: Colomn Number To Write To
    :param row:    Row Number To Write To
    :param data:   Data To Write
    """

    currentWorkSheet.write_row(colomn, row, data)

def addFormattedData(dictionary) :

    formattedData.append(dictionary)
