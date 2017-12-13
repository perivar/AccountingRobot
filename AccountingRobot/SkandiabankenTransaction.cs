using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace AccountingRobot
{
    public class SkandiabankenTransaction
    {
        public enum AccountingTypeEnum
        {
            CostOfGoods,
            CostForReselling,
            CostOfAdvertising,
            CostOfWebShop,
            CostOfDomain,
            CostOfServer,
            TransferIncome
        };

        private static Regex purchasePattern = new Regex(@"(\*0463)\s(\d+\.\d+)\s(\w+)\s(\d+\.\d+)\s([\w\.\*\s]+)\s(Kurs\:)\s(\d+\.\d+)", RegexOptions.Compiled);
        private static Regex transferPattern = new Regex(@"Fra\:\s([\w\s]+)\sBetalt\:\s(\d+\.\d+.\d+)", RegexOptions.Compiled);

        public DateTime TransactionDate { get; set; }
        public DateTime InterestDate { get; set; }
        public long ArchiveReference { get; set; }
        public string Type { get; set; }
        public string Text { get; set; }
        public decimal OutAccount { get; set; }
        public decimal InAccount { get; set; }
        public decimal AccountChange { get; set; }
        public AccountingTypeEnum AccountingType { get; set; }

        public override string ToString()
        {
            return string.Format("{0:yyyy-MM-dd} {1:yyyy-MM-dd} {2} {3} {4} {5:C}", TransactionDate, InterestDate, ArchiveReference, Type, Text, AccountChange);
        }

        public string GuessAccountType()
        {
            // https://regex101.com/
            var matchPurchase = purchasePattern.Match(Text);
            if (matchPurchase.Success)
            {
                var dayAndMonth = matchPurchase.Groups[2];
                var currency = matchPurchase.Groups[3];
                var amount = matchPurchase.Groups[4];
                var vendor = matchPurchase.Groups[5].ToString();
                var exchangeRate = matchPurchase.Groups[7];

                // fix date
                int year = DateTime.Now.Year;
                var dateString = string.Format("{0}.{1}", dayAndMonth, year);
                DateTime date = DateTime.ParseExact(dateString, "dd.MM.yyyy", CultureInfo.InvariantCulture);

                if (vendor.CaseInsensitiveContains("Wazalo")
                    || vendor.CaseInsensitiveContains("Shopifycomc"))
                {
                    this.AccountingType = AccountingTypeEnum.CostOfWebShop;
                }
                else if (vendor.CaseInsensitiveContains("Facebk"))
                {
                    this.AccountingType = AccountingTypeEnum.CostOfAdvertising;
                }
                else if (vendor.CaseInsensitiveContains("Gandi"))
                {
                    this.AccountingType = AccountingTypeEnum.CostOfDomain;
                }
                else if (vendor.CaseInsensitiveContains("Scaleway"))
                {
                    this.AccountingType = AccountingTypeEnum.CostOfServer;
                }
                else
                {
                    // this could be both CostForReselling and CostOfGoods
                    this.AccountingType = AccountingTypeEnum.CostForReselling;
                }
                return string.Format("{0} {1:dd.MM.yyyy} {2} {3} {4} {5} {6:c}", GetAccountingTypeString(), date, currency, amount, vendor, exchangeRate, AccountChange);
            }

            var matchTransfer = transferPattern.Match(Text);
            if (matchTransfer.Success)
            {
                this.AccountingType = AccountingTypeEnum.TransferIncome;
                var vendor = matchTransfer.Groups[1];
                var date = matchTransfer.Groups[2];
                return string.Format("{0} {1} {2} {3:c}", GetAccountingTypeString(), vendor, date, AccountChange);
            }

            /*
            if (Text.CaseInsensitiveContains("Aliexpress"))
            {
                // *0463 06.12 Usd 20.87 Www.Aliexpress.Com Kurs: 8.4686
            }
            else if (Text.CaseInsensitiveContains("The Currency Cloud"))
            {
                // Fra: The Currency Cloud Ltd Betalt: 08.12.17
            }
            else if (Text.CaseInsensitiveContains("Paypal Pte Ltd"))
            {
                // Fra: Boa Re Paypal Pte Ltd Betalt: 01.12.17
            }
            else if (Text.CaseInsensitiveContains("Gandi Net"))
            {
                // *0463 04.12 Nok 143.38 Gandi Net Kurs: 1.0000
            }
            else if (Text.CaseInsensitiveContains("Scaleway"))
            {
                // *0463 01.12 Eur 2.99 Scaleway Kurs: 10.1271
            }
            else if (Text.CaseInsensitiveContains("Facebk"))
            {
                // *0463 01.12 Nok 261.80 Facebk* Eg68leazj2 Kurs: 1.0000
                // *0463 31.10 Usd 18.67 Facebk C9gfde2hd2 Kurs: 8.3674
            }
            else if (Text.CaseInsensitiveContains("Shopifycomc"))
            {
                // *0463 14.11 Usd 80.08 41228928 Shopifycomc Kurs: 8.4366
            }
            */

            if (AccountChange > 0)
            {
                return string.Format("OVERFØRSEL {0} {1:c}", Text, AccountChange);
            } else
            {
                return string.Format("KJØP {0} {1:c}", Text, AccountChange);
            }
        }

        private string GetAccountingTypeString()
        {
            var accountingTypeString = "";
            switch (this.AccountingType)
            {
                case AccountingTypeEnum.CostOfGoods:
                    accountingTypeString = "VAREKOSTNAD";
                    break;
                case AccountingTypeEnum.CostForReselling:
                    accountingTypeString = "FORBRUK FOR VIDERESALG";
                    break;
                case AccountingTypeEnum.CostOfAdvertising:
                    accountingTypeString = "REKLAME KOSTNADER";
                    break;
                case AccountingTypeEnum.CostOfWebShop:
                    accountingTypeString = "NETTBUTIKK KOSTNADER";
                    break;
                case AccountingTypeEnum.CostOfDomain:
                    accountingTypeString = "DOMENE KOSTNADER";
                    break;
                case AccountingTypeEnum.CostOfServer:
                    accountingTypeString = "SERVER KOSTNADER";
                    break;
                case AccountingTypeEnum.TransferIncome:
                    accountingTypeString = "OVERFØRSEL";
                    break;
            }

            return accountingTypeString;
        }
    }
}
