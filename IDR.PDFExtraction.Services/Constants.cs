using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IDR.PDFExtraction.Services
{
    public class Constants
    {
        public static readonly string FOLDER_NAME = @"C:\Users\aogorfinadmin\Desktop\IDR";
        public static readonly string EXCEL_FILE_NAME = @"C:\Users\aogorfinadmin\Desktop\IDR.xlsx";

        public static readonly string RULE1 = "Place of Service Code:";
        public static readonly string RULE2 = "Date of the qualified IDR item or service:";
        public static readonly string RULE3 = "Dispute Reference Number:";

        public static List<string> GetRules()
        {
            List<string> rules = new()
            {
                RULE1,
                RULE2,
                RULE3
            };
            return rules;
        }
    }
}
