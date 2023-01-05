using IDR.PDFExtraction.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace IDR.PDFExtraction.Main
{
    internal class Program
    {
        static void Main(string[] args)
        {
            FileService fileService = new();
            ExcelService excelService = new();

            //get files
            Console.WriteLine("Starting PDF data extraction...");
            FileInfo[] files = fileService.GetFiles(Constants.FOLDER_NAME);

            //Initilize ExcelService variables
            excelService.ExcelFilePath = Constants.EXCEL_FILE_NAME;
            excelService.ExcelApplication = new Excel.Application();
            excelService.ExcelWorkBook = excelService.ExcelApplication.Workbooks.Add(System.Reflection.Missing.Value);
            excelService.ExcelWorkSheet = excelService.ExcelApplication.Worksheets.Add(System.Reflection.Missing.Value);

            excelService.OpenExcel();

            //Loop through each file and extract data
            foreach (var file in files)
            {
                Console.WriteLine($"Processing document - {file.Name}");
                var lineItems = fileService.ExtractTextFromPDF(file.FullName);
                if (lineItems != null)
                {
                    Console.WriteLine("Text extracted successfully.");
                    Console.WriteLine("Writing data to file....");

                    foreach (var claim in lineItems.Claims)
                    {
                        //Write extracted data to Excel
                        excelService.AddDataToExcel(lineItems.DisputeReferenceNumber, claim.ClaimTitle, claim.ClaimNumber, claim.ServiceCode, lineItems.NoticeDate);
                    }
                    Console.WriteLine($"Added {lineItems.Claims.Count} claim(s)");
                }
            }
            excelService.CloseExcel();
        }
    }

}
