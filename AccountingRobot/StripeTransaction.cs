using System;

namespace AccountingRobot
{
    public class StripeTransaction
    {
        public string TransactionID { get; set; }
        public DateTime Created { get; set; }
        public DateTime AvailableOn { get; set; }
        public bool Paid { get; set; }
        public string CustomerEmail { get; set; }
        public decimal Amount { get; set; }
        public decimal Net { get; set; }
        public decimal Fee { get; set; }
        public string Currency { get; set; }
        public string Description { get; set; }
        public string Status { get; set; }
    }
}
