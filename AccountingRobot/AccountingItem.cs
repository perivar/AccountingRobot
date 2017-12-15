using System;
using System.Collections.Generic;
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
        public string Text { get; set; }
        public string Gateway { get; set; }

        public string NumSale { get; set; }
        public string NumPurchase { get; set; }

        public decimal PurchaseUSD { get; set; }
        public decimal AmountNOK { get; set; }

        public decimal AccountPaypal { get; set; }
        public decimal AccountStripe { get; set; }
        public decimal AccountBank { get; set; }

        public decimal VATPurchase { get; set; }
        public decimal VATSales { get; set; }

        public decimal SalesVAT { get; set; }
        public decimal SalesExlVAT { get; set; }

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
            Map(m => m.Date).TypeConverterOption.Format("dd.MM.yyyy");
            Map(m => m.Number);
            Map(m => m.Text);
            Map(m => m.Gateway);
            Map(m => m.NumSale);
            Map(m => m.NumPurchase);
            Map(m => m.PurchaseUSD);
            Map(m => m.AmountNOK);

            Map(m => m.AccountPaypal);
            Map(m => m.AccountStripe);
            Map(m => m.AccountBank);

            Map(m => m.VATPurchase);
            Map(m => m.VATSales);

            Map(m => m.SalesVAT); 
            Map(m => m.SalesExlVAT);

            Map(m => m.CostOfGoods);
            Map(m => m.CostForReselling);
            Map(m => m.CostOfData);
            Map(m => m.CostOfPhoneInternet);
            Map(m => m.CostOfAdvertising);
            Map(m => m.CostOfOther);
    }
}
}
