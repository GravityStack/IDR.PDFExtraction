using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using System.Text;
using System.Text.RegularExpressions;

namespace IDR.PDFExtraction.Services
{
    public class FileService
    {
        /// <summary>
        ///Get files to process 
        /// </summary>
        /// <param name="folder"></param>
        /// <returns></returns>
        public FileInfo[] GetFiles(string folder)
        {
            DirectoryInfo di = new DirectoryInfo(folder);
            return di.GetFiles();
        }

        /// <summary>
        /// Extract text and return a Line Item object
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public LineItem ExtractTextFromPDF(string filePath)
        {
            PdfReader pdfReader = new(filePath);
            PdfDocument pdfDoc = new(pdfReader);
            StringBuilder pageContent = new StringBuilder();

            for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
            {
                ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                pageContent.Append(PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), strategy));
            }

            LineItem lineItem = ProcessExtractedText(pageContent);

            pdfDoc.Close();
            pdfReader.Close();

            return lineItem;
        }
        
        private LineItem ProcessExtractedText(StringBuilder pageContent)
        {
            StringBuilder claimTitle = new();
            StringBuilder claimNumber = new();
            string noticeDate = String.Empty;
            string disputeReferenceNumber = String.Empty;
            StringBuilder serviceCode = new();

            string[] pdfData = pageContent.ToString().Split('\n');

            //Number of Claims
            int numberOfClaims = Regex.Matches(pageContent.ToString(), "Claim Number").Count;

            for (int i = 0; i < pdfData.Length; i++)
            {
                if (pdfData[i].Contains("Dispute Reference Number"))
                {
                    disputeReferenceNumber = MiscTextFixes(pdfData[i]);
                }
                if (pdfData[i].Contains("Claim Number"))
                {
                    claimNumber.Append($"{pdfData[i + 1]} _");
                    claimTitle.Append($"{pdfData[i - 1]} _");
                }
                if (pdfData[i] == "Service Code:")
                {
                    serviceCode.Append($"{pdfData[i + 1]} _");
                }
            }
            noticeDate = pageContent.ToString().Substring(pageContent.Length - 10);

            return new LineItem
            {
                DisputeReferenceNumber = disputeReferenceNumber,
                Claims = GetClaims(numberOfClaims, claimTitle, claimNumber, serviceCode),
                NoticeDate = noticeDate
            };
        }
       
        /// <summary>
        /// Create a list of Claims objects
        /// </summary>
        /// <param name="numberOfClaims"></param>
        /// <param name="claimNumber"></param>
        /// <param name="claimTitle"></param>
        /// <param name="serviceCode"></param>
        /// <returns></returns>
        private List<Claim> GetClaims(int numberOfClaims, StringBuilder claimTitle, StringBuilder claimNumber, StringBuilder serviceCode)
        {
            List<Claim> claimList = new List<Claim>();
            string fixedValue = string.Empty;

            for (int i = 0; i < numberOfClaims; i++)
            {
                Claim claim = new()
                {
                    ClaimNumber = MiscTextFixes(claimNumber.ToString().Split("_")[i]),
                    ClaimTitle = claimTitle.ToString().Split("_")[i],
                    ServiceCode = MiscTextFixes(serviceCode.ToString().Split("_")[i])
                };
                claimList.Add(claim);
            }
            return claimList;
        }
        
        private static string MiscTextFixes(string stringToFix)
        {
            foreach (string? item in Constants.GetRules())
            {
                if (stringToFix.Contains(item))
                    return stringToFix.Replace(item, "");
            }
            return stringToFix;
        }
    }
}