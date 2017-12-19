using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccountingRobot
{
    public class PayPalTransaction
    {
        public string TransactionID { get; set; }
        public DateTime Timestamp { get; set; }
        public string Status { get; set; }
        public string Type { get; set; }
        public decimal GrossAmount { get; set; }
        public decimal NetAmount { get; set; }
        public decimal FeeAmount { get; set; }
        public string Payer { get; set; }
        public string PayerDisplayName { get; set; }
    }
}
