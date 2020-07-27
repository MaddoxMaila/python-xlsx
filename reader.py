from xlrd import open_workbook
import xlrd
import datetime
import Include.writer as xwriter

class Reader:

    def __init__(self):

        '''
            File Name Of The .XLSX File To Open... Input File
        '''
        self.filename = None

        '''
            Hold All Strings To Be Used As Sheets In The Output File
        '''
        self.sheets_to_new_file = []

        '''
            Instance Of The Opened Input File
        '''
        self.work_book = None

        '''
            Instance Of The Sheet In The Opened Input File
        '''
        self.work_sheet = None

    '''
        Method To Open The Specified File
        Create The WorkBook & WorkSheet Instance
    '''
    def add_xfile(self, filename):

        """
        :param filename: Path To The File You Want Open
        """

        '''
            Exception Handling For IO Operation
        '''
        try:

            '''
                Create WorkBook/ Open The Input File
            '''
            self.work_book = open_workbook(filename)

            '''
                Get The First WorkSheet
            '''
            self.work_sheet = self.work_book.sheet_by_index(0)


        except IOError:

            '''
                Catch All IO Error
            '''
            print(IOError.filename)

    '''
        Get All Strings That Will Be Sheets In Output File
    '''
    def add_xtract_worksheets(self):


        '''
            All Sheets To Be Extracted Are On This Row
        '''
        sheetsRowIndex = 6

        '''
            Get The Output File WorkBook
        '''
        self.writer = xwriter.getWorkBook('./new files/output.xlsx')

        '''
            Iterate Through Only On The 6th Row
            From Column 4, It's Headers That Will Be Work Sheets On The New Structured File 
        '''
        for columnIndex in range(self.work_sheet.ncols):

            '''
                Sheets To Be Extracted Are From The 4th Column
            '''
            if columnIndex > 3:

                '''
                    Extracted Sheet String From Input File
                    It Will Be Used As A WorkBook Sheet In The Output Sheet
                '''
                sheetFromCell = self.work_sheet.cell(sheetsRowIndex, columnIndex).value

                '''
                    Process Sheet Strings To Make The Them WorkBook Sheets In Output File
                '''
                self.processSheetsExtracted(sheetFromCell, columnIndex)

        '''
            Close The Output File Workbook To Save Changes!
        '''
        xwriter.workBookClose()

        print("***************************************************************************\n*\n*\n*\t\t\t\t\t\t\tEXTRACTION PROCESS DONE\n*\n*\n*")
        print("***************************************************************************")

    '''
        Method To Start The Extraction Process Of The Sheet Strings
    '''
    def processSheetsExtracted(self, Sheet, columnIndex) :

        '''
        :param Sheet: Extracted Sheet
        :return: none
        '''

        '''
            Add Sheet To All Extracted Sheets List
        '''
        self.sheets_to_new_file.append(Sheet)

        '''
            Iterate Through All Rows Of The WorkSheet
        '''
        for rowIndex in range(self.work_sheet.nrows) :

            '''
                Data For All Sheets Start From Row 7
            '''
            if rowIndex > 6 :

                '''
                    Extracted Data
                '''
                dateFromCell    = self.work_sheet.cell(rowIndex, 1).value
                paymentFromCell = self.work_sheet.cell(rowIndex, columnIndex).value

                '''
                    Format Extracted Data
                '''
                self.formatExtractedData(Sheet, dateFromCell, paymentFromCell)

        '''
            Creates Sheet In Output File
        '''
        xwriter.addWorkSheet(Sheet)

    '''
        Method To Format All Extracted Data
    '''
    def formatExtractedData(self, sheetName, rawDate, payment) :

        '''
        :param sheetName:
        :param rawDate:
        :param payment:
        :return: None
        '''

        '''
            Destructure The Date
        '''
        y, m, d, h, m, s = xlrd.xldate_as_tuple(rawDate, self.work_book.datemode)

        date          = "{}/{}/{} {}:{}".format(y, m, d, h, m)

        '''
            Format The Extracted Date In Dictionary For Key-Value, Ease Of Transmission 
        '''
        dataSet       = {

            "sheet_name"    : sheetName,
            "date"          : date,
            "payment"       : payment

        }

        '''
            Add Data Set To Writer For Writing In The Output File
        '''
        xwriter.addFormattedData(dataSet)

'''
    Create Reader Object To Start The Extraction Process
'''
read = Reader()

'''
    Add File
'''
read.add_xfile('./original file/Input_Nominal_PRO (Pasted Values)_T (1).xlsx')

'''
    Begin With The Extraction Of Sheets Process
'''
read.add_xtract_worksheets()