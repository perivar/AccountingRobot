using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccountingRobot
{
    public class StripeTransaction
    {
        public string TransactionID { get; set; }
        public DateTime Created { get; set; }
        public bool Paid { get; set; }
        public string CustomerEmail { get; set; }
        public decimal Amount { get; set; }
        public decimal Net { get; set; }
        public decimal Fee { get; set; }
    }
}
