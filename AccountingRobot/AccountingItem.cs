using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper.Configuration;

namespace AccountingRobot
{
    public class AccountingItem
    {
        public int Periode => Date.Month;
        public DateTime Date { get; set; }

        public int Number { get; set; }
        public long ArchiveReference { get; set; }
        public string Type { get; set; } // Overføring (intern), Overførsel (ekstern), Visa, Avgift
        public string AccountingType { get; set; }
        public string Text { get; set; }

        public string Gateway { get; set; }

        public string NumSale { get; set; }
        public string NumPurchase { get; set; }

        public decimal PurchaseOtherCurrency { get; set; }
        public string OtherCurrency { get; set; }

        public decimal AccountPaypal { get; set; }
        public decimal AccountStripe { get; set; }
        public decimal AccountBank { get; set; }

        public decimal VATPurchase { get; set; }
        public decimal VATSales { get; set; }

        public decimal SalesVAT { get; set; }
        public decimal SalesVATExempt { get; set; }

        public decimal CostOfGoods { get; set; }
        public decimal CostForReselling { get; set; }
        public decimal CostOfData { get; set; }
        public decimal CostOfPhoneInternet { get; set; }
        public decimal CostOfAdvertising { get; set; }
        public decimal CostOfOther { get; set; }

        public decimal FeesBank { get; set; }
        public decimal FeesPaypal { get; set; }
        public decimal FeesStripe { get; set; }

        public decimal IncomeFinance { get; set; }
        public decimal CostOfFinance { get; set; }
    }

    public sealed class AccountingItemCsvMap : ClassMap<AccountingItem>
    {
        public AccountingItemCsvMap()
        {
            Map(m => m.Periode);
            Map(m => m.Date).TypeConverterOption.Format("yyyy.MM.dd");

            Map(m => m.Number);
            Map(m => m.ArchiveReference);
            Map(m => m.Type);
            Map(m => m.AccountingType);
            Map(m => m.Text);

            Map(m => m.Gateway);
            Map(m => m.NumSale);
            Map(m => m.NumPurchase);
            Map(m => m.PurchaseOtherCurrency);
            Map(m => m.OtherCurrency);

            Map(m => m.AccountPaypal);
            Map(m => m.AccountStripe);
            Map(m => m.AccountBank);

            Map(m => m.VATPurchase);
            Map(m => m.VATSales);

            Map(m => m.SalesVAT);
            Map(m => m.SalesVATExempt);

            Map(m => m.CostOfGoods);
            Map(m => m.CostForReselling);
            Map(m => m.CostOfData);
            Map(m => m.CostOfPhoneInternet);
            Map(m => m.CostOfAdvertising);
            Map(m => m.CostOfOther);
        }
    }
}
