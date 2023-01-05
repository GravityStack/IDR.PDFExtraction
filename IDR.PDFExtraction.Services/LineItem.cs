using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IDR.PDFExtraction.Services
{
    public class LineItem
    {
        public String? DisputeReferenceNumber { get; set; }
        public List<Claim>? Claims{ get; set; }
        public String? NoticeDate { get; set; }
    }
}
