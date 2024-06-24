using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestEpplusExel
{
    public class PreAprovedExcelModel
    {
        public string CIFCode { get; set; }
        public string IdentityNo { get; set; }
        public string Phone { get; set; }
        public string FullName { get; set; }
        public string Birthday { get; set; }
        public string AssetCode { get; set; }
        public string CategoryCode { get; set; }
        public string Type { get; set; } = string.Empty;
        public string InterestPaymentDay { get; set; }
        public string ShopCode { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public string BankAccount { get; set; }
        public string ValidResult { get; set; } = string.Empty;
    }
}
