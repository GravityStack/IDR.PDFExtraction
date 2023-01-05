using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace IDR.PDFExtraction.Services
{
    public class ExcelService
    {
        private string excelFilePath = String.Empty;
        private int rowNumber = 2; // define first row number to enter data in excel
        private Excel.Workbook myExcelWorkbook;
        private Excel.Worksheet myExcelWorksheet;
        private Excel.Application myExcelApplication;

        public string ExcelFilePath
        {
            get { return excelFilePath; }
            set { excelFilePath = value; }
        }        
        public Excel.Application ExcelApplication
        {
            get { return myExcelApplication; }
            set { myExcelApplication = value; }
        }
        public Excel.Workbook ExcelWorkBook
        {
            get { return myExcelWorkbook; }
            set { myExcelWorkbook = value; }
        }        
        public Excel.Worksheet ExcelWorkSheet
        {
            get { return myExcelWorksheet; }
            set { myExcelWorksheet = value; }
        }
        public int Rownumber
        {
            get { return rowNumber; }
            set { rowNumber = value; }
        }
        public void OpenExcel()
        {
            myExcelWorkbook = (Excel.Workbook)ExcelApplication.Workbooks.Open(excelFilePath); // open the existing excel file
            myExcelWorksheet = (Excel.Worksheet)myExcelWorkbook.Worksheets[1]; // define in which worksheet, do you want to add data
        }
        public void AddDataToExcel(string disputeReferenceNumber, string claimTitle, string claimNumber, string serviceCode, string noticeDate)
        {
            myExcelWorksheet.Cells[rowNumber, "A"] = disputeReferenceNumber;
            myExcelWorksheet.Cells[rowNumber, "B"] = claimTitle;
            myExcelWorksheet.Cells[rowNumber, "C"] = claimNumber;
            myExcelWorksheet.Cells[rowNumber, "D"] = serviceCode;
            myExcelWorksheet.Cells[rowNumber, "E"] = noticeDate;
            rowNumber++;
        }
        public void CloseExcel()
        {
            try
            {
                myExcelWorkbook.Save(); // Save data in excel
                myExcelWorkbook.Close(true, excelFilePath, System.Reflection.Missing.Value); // close the worksheet
            }
            finally
            {
                if (myExcelApplication != null)
                {
                    myExcelApplication.Quit(); // close the excel application
                }
            }
        }
    }
}
