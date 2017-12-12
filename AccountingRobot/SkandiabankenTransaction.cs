using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccountingRobot
{
    public class SkandiabankenTransaction
    {
        public DateTime TransactionDate { get; set; }
        public DateTime InterestDate { get; set; }
        public long ArchiveReference { get; set; }
        public string Type { get; set; }
        public string Text { get; set; }
        public decimal OutAccount { get; set; }
        public decimal InAccount { get; set; }
        public decimal AccountChange { get; set; }

        public override string ToString()
        {
            return string.Format("{0:yyyy-MM-dd} {1:yyyy-MM-dd} {2} {3} {4} {5:C}", TransactionDate, InterestDate, ArchiveReference, Type, Text, AccountChange);
        }
    }
}
