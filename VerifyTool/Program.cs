using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using AccountingRobot;
using ClosedXML.Excel;

namespace VerifyTool
{
    class Program
    {
        static void Main(string[] args)
        {
            // export or update accounting spreadsheet
            string accountingFileDir = ConfigurationManager.AppSettings["AccountingDir"];
            string accountingFileNamePrefix = "wazalo regnskap";
            string accountingDateFromToRegexPattern = @"(\d{4}\-\d{2}\-\d{2})\-(\d{4}\-\d{2}\-\d{2})\.xlsx$";
            var lastAccountingFileInfo = Utils.FindLastCacheFile(accountingFileDir, accountingFileNamePrefix, accountingDateFromToRegexPattern, "yyyy-MM-dd", "\\-");

            // if the cache file object has values
            if (!lastAccountingFileInfo.Equals(default(KeyValuePair<DateTime, string>)))
            {
                Console.Out.WriteLine("Found an accounting spreadsheet from {0:yyyy-MM-dd}", lastAccountingFileInfo.From);
                Console.Out.WriteLine("Checking PayPal transactions...");
                CheckPayPal(lastAccountingFileInfo.FilePath);

                Console.Out.WriteLine("Checking Sripe transactions...");
                CheckStripe(lastAccountingFileInfo.FilePath);
            }
            else
            {
                Console.Out.WriteLine("Error! No accounting spreadsheet found!.");
                return;
            }

            Console.ReadLine();
        }

        static void CheckPayPal(string filePath)
        {
            var payPalTransactions = PayPalFactory.Instance.GetLatest();
            //Console.Out.WriteLine("Successfully read PayPal transactions ...");

            var existingAccountingItems = GetSpreadsheet(filePath);
            if (existingAccountingItems.Count() > 0)
            {
                var existingPayPalTransactions =
                    (from row in existingAccountingItems
                     where
                     row.Value.Gateway == "paypal"
                     orderby row.Value.Number ascending
                     select row);

                // identify elements from the paypal list that doesn't exist in the spreadsheet
                var foundDict = new Dictionary<string, PayPalTransaction>();
                var notFoundDict = new Dictionary<string, PayPalTransaction>();
                foreach (var payPalTransaction in payPalTransactions)
                {
                    var foundPaypalTransaction =
                        (from row in existingPayPalTransactions
                         where
                         row.Value.TransactionID == payPalTransaction.TransactionID
                         orderby row.Value.Number ascending
                         select row);

                    if (foundPaypalTransaction.Count() > 0)
                    {
                        foundDict.Add(payPalTransaction.TransactionID, payPalTransaction);
                    }
                    else
                    {
                        notFoundDict.Add(payPalTransaction.TransactionID, payPalTransaction);
                    }
                }

                if (notFoundDict.Count() > 0) Console.Out.WriteLine("Number of PayPal transactions not found in accounting spreadsheet: {0}", notFoundDict.Count());
                foreach (var notFound in notFoundDict)
                {
                    Console.Out.WriteLine("{0}", notFound.Value);
                }
            }
        }

        static void CheckStripe(string filePath)
        {
            var stripeTransactions = StripeChargeFactory.Instance.GetLatest();
            //Console.Out.WriteLine("Successfully read Stripe transactions ...");

            var stripePayoutTransactions = StripePayoutFactory.Instance.GetLatest();
            //Console.Out.WriteLine("Successfully read Stripe payout transactions ...");

            // combine to one list
            stripeTransactions.AddRange(stripePayoutTransactions);

            var existingAccountingItems = GetSpreadsheet(filePath);
            if (existingAccountingItems.Count() > 0)
            {
                var existingStripeTransactions =
                    (from row in existingAccountingItems
                     where
                     row.Value.Gateway == "stripe"
                     orderby row.Value.Number ascending
                     select row);

                // identify elements from the stripe list that doesn't exist in the spreadsheet
                var foundDict = new Dictionary<string, StripeTransaction>();
                var notFoundDict = new Dictionary<string, StripeTransaction>();
                foreach (var stripeTransaction in stripeTransactions)
                {
                    var foundStripeTransaction =
                        (from row in existingStripeTransactions
                         where
                         row.Value.TransactionID == stripeTransaction.TransactionID
                         orderby row.Value.Number ascending
                         select row);

                    if (foundStripeTransaction.Count() > 0)
                    {
                        foundDict.Add(stripeTransaction.TransactionID, stripeTransaction);
                    }
                    else
                    {
                        notFoundDict.Add(stripeTransaction.TransactionID, stripeTransaction);
                    }
                }

                if (notFoundDict.Count() > 0) Console.Out.WriteLine("Number of Stripe transactions not found in accounting spreadsheet: {0}", notFoundDict.Count());
                foreach (var notFound in notFoundDict)
                {
                    Console.Out.WriteLine("{0}", notFound.Value);
                }
            }
        }

        static Dictionary<IXLTableRow, AccountingItem> GetSpreadsheet(string filePath)
        {
            XLWorkbook wb = new XLWorkbook(filePath);
            IXLWorksheet ws = wb.Worksheet("Bilagsjournal");

            IXLTables tables = ws.Tables;
            IXLTable table = tables.FirstOrDefault();

            var existingAccountingItems = new Dictionary<IXLTableRow, AccountingItem>();
            if (table != null)
            {
                foreach (var row in table.DataRange.Rows())
                {
                    var accountingItem = new AccountingItem();
                    accountingItem.Date = ExcelUtils.GetExcelField<DateTime>(row, "Dato");
                    accountingItem.Number = ExcelUtils.GetExcelField<int>(row, "Bilagsnr.");
                    accountingItem.ArchiveReference = ExcelUtils.GetExcelField<long>(row, "Arkivreferanse").ToString();
                    accountingItem.TransactionID = ExcelUtils.GetExcelField<string>(row, "TransaksjonsId");
                    accountingItem.Type = ExcelUtils.GetExcelField<string>(row, "Type");
                    accountingItem.AccountingType = ExcelUtils.GetExcelField<string>(row, "Regnskapstype");
                    accountingItem.Text = ExcelUtils.GetExcelField<string>(row, "Tekst");
                    accountingItem.CustomerName = ExcelUtils.GetExcelField<string>(row, "Kundenavn");
                    accountingItem.ErrorMessage = ExcelUtils.GetExcelField<string>(row, "Feilmelding");
                    accountingItem.Gateway = ExcelUtils.GetExcelField<string>(row, "Gateway");
                    accountingItem.NumSale = ExcelUtils.GetExcelField<string>(row, "Num Salg");
                    accountingItem.NumPurchase = ExcelUtils.GetExcelField<string>(row, "Num Kjøp");
                    accountingItem.PurchaseOtherCurrency = ExcelUtils.GetExcelField<decimal>(row, "Kjøp annen valuta");
                    accountingItem.OtherCurrency = ExcelUtils.GetExcelField<string>(row, "Annen valuta");

                    accountingItem.AccountPaypal = ExcelUtils.GetExcelField<decimal>(row, "Paypal");    // 1910
                    accountingItem.AccountStripe = ExcelUtils.GetExcelField<decimal>(row, "Stripe");    // 1915
                    accountingItem.AccountVipps = ExcelUtils.GetExcelField<decimal>(row, "Vipps");  // 1918
                    accountingItem.AccountBank = ExcelUtils.GetExcelField<decimal>(row, "Bank");    // 1920

                    accountingItem.VATPurchase = ExcelUtils.GetExcelField<decimal>(row, "MVA Kjøp");
                    accountingItem.VATSales = ExcelUtils.GetExcelField<decimal>(row, "MVA Salg");

                    accountingItem.SalesVAT = ExcelUtils.GetExcelField<decimal>(row, "Salg mva-pliktig");   // 3000
                    accountingItem.SalesVATExempt = ExcelUtils.GetExcelField<decimal>(row, "Salg avgiftsfritt");    // 3100

                    accountingItem.CostOfGoods = ExcelUtils.GetExcelField<decimal>(row, "Varekostnad"); // 4005
                    accountingItem.CostForReselling = ExcelUtils.GetExcelField<decimal>(row, "Forbruk for videresalg"); // 4300
                    accountingItem.CostForSalary = ExcelUtils.GetExcelField<decimal>(row, "Lønn");  // 5000
                    accountingItem.CostForSalaryTax = ExcelUtils.GetExcelField<decimal>(row, "Arb.giver avgift");   // 5400
                    accountingItem.CostForDepreciation = ExcelUtils.GetExcelField<decimal>(row, "Avskrivninger");   // 6000
                    accountingItem.CostForShipping = ExcelUtils.GetExcelField<decimal>(row, "Frakt");   // 6100
                    accountingItem.CostForElectricity = ExcelUtils.GetExcelField<decimal>(row, "Strøm");    // 6340 
                    accountingItem.CostForToolsInventory = ExcelUtils.GetExcelField<decimal>(row, "Verktøy inventar");  // 6500
                    accountingItem.CostForMaintenance = ExcelUtils.GetExcelField<decimal>(row, "Vedlikehold");  // 6695
                    accountingItem.CostForFacilities = ExcelUtils.GetExcelField<decimal>(row, "Kontorkostnader");   // 6800 

                    accountingItem.CostOfData = ExcelUtils.GetExcelField<decimal>(row, "Datakostnader");    // 6810 
                    accountingItem.CostOfPhoneInternet = ExcelUtils.GetExcelField<decimal>(row, "Telefon Internett");   // 6900
                    accountingItem.CostForTravelAndAllowance = ExcelUtils.GetExcelField<decimal>(row, "Reise og Diett");    // 7140
                    accountingItem.CostOfAdvertising = ExcelUtils.GetExcelField<decimal>(row, "Reklamekostnader");  // 7330
                    accountingItem.CostOfOther = ExcelUtils.GetExcelField<decimal>(row, "Diverse annet");   // 7700

                    accountingItem.FeesBank = ExcelUtils.GetExcelField<decimal>(row, "Gebyrer Bank");   // 7770
                    accountingItem.FeesPaypal = ExcelUtils.GetExcelField<decimal>(row, "Gebyrer Paypal");   // 7780
                    accountingItem.FeesStripe = ExcelUtils.GetExcelField<decimal>(row, "Gebyrer Stripe");   // 7785 

                    accountingItem.CostForEstablishment = ExcelUtils.GetExcelField<decimal>(row, "Etableringskostnader");   // 7790

                    accountingItem.IncomeFinance = ExcelUtils.GetExcelField<decimal>(row, "Finansinntekter");   // 8099
                    accountingItem.CostOfFinance = ExcelUtils.GetExcelField<decimal>(row, "Finanskostnader");   // 8199

                    accountingItem.Investments = ExcelUtils.GetExcelField<decimal>(row, "Investeringer");   // 1200
                    accountingItem.AccountsReceivable = ExcelUtils.GetExcelField<decimal>(row, "Kundefordringer");  // 1500
                    accountingItem.PersonalWithdrawal = ExcelUtils.GetExcelField<decimal>(row, "Privat uttak");
                    accountingItem.PersonalDeposit = ExcelUtils.GetExcelField<decimal>(row, "Privat innskudd");

                    existingAccountingItems.Add(row, accountingItem);
                }
            }
            return existingAccountingItems;
        }
    }
}
