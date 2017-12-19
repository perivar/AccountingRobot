using System;
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
        public string CustomerName { get; set; }
        public string ErrorMessage { get; set; }

        public string Gateway { get; set; }

        public string NumSale { get; set; }
        public string NumPurchase { get; set; }

        public decimal PurchaseOtherCurrency { get; set; }
        public string OtherCurrency { get; set; }

        public decimal AccountPaypal { get; set; }  // 1910
        public decimal AccountStripe { get; set; }  // 1915
        public decimal AccountVipps { get; set; }   // 1918
        public decimal AccountBank { get; set; }    // 1920

        public decimal VATPurchase { get; set; }
        public decimal VATSales { get; set; }

        public decimal SalesVAT { get; set; }       // 3000
        public decimal SalesVATExempt { get; set; } // 3100

        public decimal CostOfGoods { get; set; }            // 4005
        public decimal CostForReselling { get; set; }       // 4300
        public decimal CostForSalary { get; set; }          // 5000
        public decimal CostForSalaryTax { get; set; }       // 5400
        public decimal CostForDepreciation { get; set; }    // 6000
        public decimal CostForShipping { get; set; }        // 6100
        public decimal CostForElectricity { get; set; }     // 6340 
        public decimal CostForToolsInventory { get; set; }      // 6500
        public decimal CostForMaintenance { get; set; }         // 6695
        public decimal CostForFacilities { get; set; }          // 6800 

        public decimal CostOfData { get; set; }                 // 6810 
        public decimal CostOfPhoneInternet { get; set; }        // 6900
        public decimal CostForTravelAndAllowance { get; set; }  // 7140
        public decimal CostOfAdvertising { get; set; }          // 7330
        public decimal CostOfOther { get; set; }                // 7700

        public decimal FeesBank { get; set; }                   // 7770
        public decimal FeesPaypal { get; set; }                 // 7780
        public decimal FeesStripe { get; set; }                 // 7785 

        public decimal CostForEstablishment { get; set; }       // 7790

        public decimal IncomeFinance { get; set; }              // 8099
        public decimal CostOfFinance { get; set; }              // 8199
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
            Map(m => m.CustomerName);
            Map(m => m.ErrorMessage);

            Map(m => m.Gateway);
            Map(m => m.NumSale);
            Map(m => m.NumPurchase);
            Map(m => m.PurchaseOtherCurrency);
            Map(m => m.OtherCurrency);

            Map(m => m.AccountPaypal);
            Map(m => m.AccountStripe);
            Map(m => m.AccountVipps);
            Map(m => m.AccountBank);

            Map(m => m.VATPurchase);
            Map(m => m.VATSales);

            Map(m => m.SalesVAT);
            Map(m => m.SalesVATExempt);

            Map(m => m.CostOfGoods);
            Map(m => m.CostForReselling);
            Map(m => m.CostForSalary);
            Map(m => m.CostForSalaryTax);
            Map(m => m.CostForDepreciation);
            Map(m => m.CostForShipping);
            Map(m => m.CostForElectricity);
            Map(m => m.CostForToolsInventory);
            Map(m => m.CostForMaintenance);
            Map(m => m.CostForFacilities);

            Map(m => m.CostOfData);
            Map(m => m.CostOfPhoneInternet);
            Map(m => m.CostForTravelAndAllowance);
            Map(m => m.CostOfAdvertising);
            Map(m => m.CostOfOther);

            Map(m => m.FeesBank);
            Map(m => m.FeesPaypal);
            Map(m => m.FeesStripe);

            Map(m => m.CostForEstablishment);

            Map(m => m.IncomeFinance);
            Map(m => m.CostOfFinance);
        }
    }
}
