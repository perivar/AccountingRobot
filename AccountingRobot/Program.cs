using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using AliOrderScraper;
using OberloScraper;
using ClosedXML.Excel;
using System.Data;

namespace AccountingRobot
{
    partial class Program
    {
        static void Main(string[] args)
        {
            // process the transactions and create accounting overview
            var customerNames = new List<string>();
            var accountingShopifyItems = ProcessShopifyStatement(customerNames);

            // select only distinct 
            customerNames = customerNames.Distinct().ToList();

            // find latest skandiabanken transaction spreadsheet
            var accountingBankItems = default(List<AccountingItem>);
            string cacheDir = ConfigurationManager.AppSettings["CacheDir"];
            string cacheFileNamePrefix = "97132735232";
            string dateFromToRegexPattern = @"(\d{4}_\d{2}_\d{2})\-(\d{4}_\d{2}_\d{2})\.xlsx$";
            var lastCacheFile = Utils.FindLastCacheFile(cacheDir, cacheFileNamePrefix, dateFromToRegexPattern, "yyyy_MM_dd", "_");

            // if the cache file object has values
            if (!lastCacheFile.Equals(default(KeyValuePair<DateTime, string>)))
            {
                accountingBankItems = ProcessBankAccountStatement(lastCacheFile.Value, customerNames);
            }
            else
            {
                Console.Out.WriteLine("Error! No SBanken transaction file found!.");
                return;
                //string skandiabankenXLSX = @"C:\Users\pnerseth\Amazon Drive\Documents\Private\wazalo\regnskap\97132735232_2017_01_01-2017_12_20.xlsx";
                //accountingBankItems = ProcessBankAccountStatement(skandiabankenXLSX);
            }

            // merge into one list
            accountingShopifyItems.AddRange(accountingBankItems);

            // and sort (by ascending)
            var accountingItems = accountingShopifyItems.OrderBy(o => o.Date).ToList();

            // export to excel file
            var now = DateTime.Now;
            var fileName = string.Format("Accounting {0:yyyy-MM-dd}.xlsx", now);
            ExportToExcel(fileName, accountingItems);

            Console.ReadLine();
        }

        static void ExportToExcel(string filePath, List<AccountingItem> accountingItems)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Control", typeof(String));

            dt.Columns.Add("Periode", typeof(int));
            dt.Columns.Add("Date", typeof(DateTime));
            dt.Columns.Add("Number", typeof(int));
            dt.Columns.Add("ArchiveReference", typeof(long));
            dt.Columns.Add("Type", typeof(string));
            dt.Columns.Add("AccountingType", typeof(string));
            dt.Columns.Add("Text", typeof(string));
            dt.Columns.Add("CustomerName", typeof(string));
            dt.Columns.Add("ErrorMessage", typeof(string));
            dt.Columns.Add("Gateway", typeof(string));
            dt.Columns.Add("NumSale", typeof(string));
            dt.Columns.Add("NumPurchase", typeof(string));
            dt.Columns.Add("PurchaseOtherCurrency", typeof(decimal));
            dt.Columns.Add("OtherCurrency", typeof(string));

            dt.Columns.Add("AccountPaypal", typeof(decimal));           // 1910
            dt.Columns.Add("AccountStripe", typeof(decimal));           // 1915
            dt.Columns.Add("AccountVipps", typeof(decimal));            // 1918
            dt.Columns.Add("AccountBank", typeof(decimal));             // 1920

            dt.Columns.Add("VATPurchase", typeof(decimal));
            dt.Columns.Add("VATSales", typeof(decimal));

            dt.Columns.Add("SalesVAT", typeof(decimal));                // 3000
            dt.Columns.Add("SalesVATExempt", typeof(decimal));          // 3100

            dt.Columns.Add("CostOfGoods", typeof(decimal));             // 4005
            dt.Columns.Add("CostForReselling", typeof(decimal));        // 4300
            dt.Columns.Add("CostForSalary", typeof(decimal));           // 5000
            dt.Columns.Add("CostForSalaryTax", typeof(decimal));        // 5400
            dt.Columns.Add("CostForDepreciation", typeof(decimal));     // 6000
            dt.Columns.Add("CostForShipping", typeof(decimal));         // 6100
            dt.Columns.Add("CostForElectricity", typeof(decimal));      // 6340 
            dt.Columns.Add("CostForToolsInventory", typeof(decimal));   // 6500
            dt.Columns.Add("CostForMaintenance", typeof(decimal));      // 6695
            dt.Columns.Add("CostForFacilities", typeof(decimal));       // 6800 

            dt.Columns.Add("CostOfData", typeof(decimal));              // 6810 
            dt.Columns.Add("CostOfPhoneInternet", typeof(decimal));     // 6900
            dt.Columns.Add("CostForTravelAndAllowance", typeof(decimal));// 7140
            dt.Columns.Add("CostOfAdvertising", typeof(decimal));       // 7330
            dt.Columns.Add("CostOfOther", typeof(decimal));             // 7700

            dt.Columns.Add("FeesBank", typeof(decimal));                // 7770
            dt.Columns.Add("FeesPaypal", typeof(decimal));              // 7780
            dt.Columns.Add("FeesStripe", typeof(decimal));              // 7785 

            dt.Columns.Add("CostForEstablishment", typeof(decimal));    // 7790

            dt.Columns.Add("IncomeFinance", typeof(decimal));           // 8099
            dt.Columns.Add("CostOfFinance", typeof(decimal));           // 8199

            dt.Columns.Add("Investments", typeof(decimal));             // 1200
            dt.Columns.Add("AccountsReceivable", typeof(decimal));      // 1500
            dt.Columns.Add("PersonalWithdrawal", typeof(decimal));
            dt.Columns.Add("PersonalDeposit", typeof(decimal));

            dt.Columns.Add("SumPreRounding", typeof(decimal));
            dt.Columns.Add("SumRounded", typeof(decimal));

            foreach (var accountingItem in accountingItems)
            {
                dt.Rows.Add(
                    "",
                    accountingItem.Periode,
                    accountingItem.Date,
                    accountingItem.Number,
                    accountingItem.ArchiveReference,
                    accountingItem.Type,
                    accountingItem.AccountingType,
                    accountingItem.Text,
                    accountingItem.CustomerName,
                    accountingItem.ErrorMessage,
                    accountingItem.Gateway,
                    accountingItem.NumSale,
                    accountingItem.NumPurchase,
                    accountingItem.PurchaseOtherCurrency,
                    accountingItem.OtherCurrency,

                    accountingItem.AccountPaypal,               // 1910
                    accountingItem.AccountStripe,               // 1915
                    accountingItem.AccountVipps,                // 1918
                    accountingItem.AccountBank,                 // 1920

                    accountingItem.VATPurchase,
                    accountingItem.VATSales,

                    accountingItem.SalesVAT,                    // 3000
                    accountingItem.SalesVATExempt,              // 3100

                    accountingItem.CostOfGoods,                 // 4005
                    accountingItem.CostForReselling,            // 4300
                    accountingItem.CostForSalary,               // 5000
                    accountingItem.CostForSalaryTax,            // 5400
                    accountingItem.CostForDepreciation,         // 6000
                    accountingItem.CostForShipping,             // 6100
                    accountingItem.CostForElectricity,          // 6340 
                    accountingItem.CostForToolsInventory,       // 6500
                    accountingItem.CostForMaintenance,          // 6695
                    accountingItem.CostForFacilities,           // 6800 

                    accountingItem.CostOfData,                  // 6810 
                    accountingItem.CostOfPhoneInternet,         // 6900
                    accountingItem.CostForTravelAndAllowance,   // 7140
                    accountingItem.CostOfAdvertising,           // 7330
                    accountingItem.CostOfOther,                 // 7700

                    accountingItem.FeesBank,                    // 7770
                    accountingItem.FeesPaypal,                  // 7780
                    accountingItem.FeesStripe,                  // 7785 

                    accountingItem.CostForEstablishment,        // 7790

                    accountingItem.IncomeFinance,               // 8099
                    accountingItem.CostOfFinance,               // 8199

                    accountingItem.Investments,                 // 1200
                    accountingItem.AccountsReceivable,          // 1500
                    accountingItem.PersonalWithdrawal,
                    accountingItem.PersonalDeposit
                    );
            }

            // Build Excel spreadsheet using Closed XML
            using (XLWorkbook wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add(dt, "Accounting");
                var table = ws.Tables.First();
                table.Theme = XLTableTheme.TableStyleMedium15;
                table.ShowTotalsRow = true;

                // set sum functions for each of the table columns 
                table.Field("AccountPaypal").TotalsRowFunction = XLTotalsRowFunction.Sum;              // 1910
                table.Field("AccountStripe").TotalsRowFunction = XLTotalsRowFunction.Sum;              // 1915
                table.Field("AccountVipps").TotalsRowFunction = XLTotalsRowFunction.Sum;               // 1918
                table.Field("AccountBank").TotalsRowFunction = XLTotalsRowFunction.Sum;                // 1920

                table.Field("VATPurchase").TotalsRowFunction = XLTotalsRowFunction.Sum;
                table.Field("VATSales").TotalsRowFunction = XLTotalsRowFunction.Sum;

                table.Field("SalesVAT").TotalsRowFunction = XLTotalsRowFunction.Sum;                   // 3000
                table.Field("SalesVATExempt").TotalsRowFunction = XLTotalsRowFunction.Sum;             // 3100

                table.Field("CostOfGoods").TotalsRowFunction = XLTotalsRowFunction.Sum;                // 4005
                table.Field("CostForReselling").TotalsRowFunction = XLTotalsRowFunction.Sum;           // 4300
                table.Field("CostForSalary").TotalsRowFunction = XLTotalsRowFunction.Sum;              // 5000
                table.Field("CostForSalaryTax").TotalsRowFunction = XLTotalsRowFunction.Sum;           // 5400
                table.Field("CostForDepreciation").TotalsRowFunction = XLTotalsRowFunction.Sum;        // 6000
                table.Field("CostForShipping").TotalsRowFunction = XLTotalsRowFunction.Sum;            // 6100
                table.Field("CostForElectricity").TotalsRowFunction = XLTotalsRowFunction.Sum;         // 6340 
                table.Field("CostForToolsInventory").TotalsRowFunction = XLTotalsRowFunction.Sum;      // 6500
                table.Field("CostForMaintenance").TotalsRowFunction = XLTotalsRowFunction.Sum;         // 6695
                table.Field("CostForFacilities").TotalsRowFunction = XLTotalsRowFunction.Sum;          // 6800 

                table.Field("CostOfData").TotalsRowFunction = XLTotalsRowFunction.Sum;                 // 6810 
                table.Field("CostOfPhoneInternet").TotalsRowFunction = XLTotalsRowFunction.Sum;        // 6900
                table.Field("CostForTravelAndAllowance").TotalsRowFunction = XLTotalsRowFunction.Sum;  // 7140
                table.Field("CostOfAdvertising").TotalsRowFunction = XLTotalsRowFunction.Sum;          // 7330
                table.Field("CostOfOther").TotalsRowFunction = XLTotalsRowFunction.Sum;                // 7700

                table.Field("FeesBank").TotalsRowFunction = XLTotalsRowFunction.Sum;                   // 7770
                table.Field("FeesPaypal").TotalsRowFunction = XLTotalsRowFunction.Sum;                 // 7780
                table.Field("FeesStripe").TotalsRowFunction = XLTotalsRowFunction.Sum;                 // 7785 

                table.Field("CostForEstablishment").TotalsRowFunction = XLTotalsRowFunction.Sum;       // 7790

                table.Field("IncomeFinance").TotalsRowFunction = XLTotalsRowFunction.Sum;              // 8099
                table.Field("CostOfFinance").TotalsRowFunction = XLTotalsRowFunction.Sum;              // 8199

                table.Field("Investments").TotalsRowFunction = XLTotalsRowFunction.Sum;                // 1200
                table.Field("AccountsReceivable").TotalsRowFunction = XLTotalsRowFunction.Sum;         // 1500
                table.Field("PersonalWithdrawal").TotalsRowFunction = XLTotalsRowFunction.Sum;
                table.Field("PersonalDeposit").TotalsRowFunction = XLTotalsRowFunction.Sum;

                // set formulas
                // get all rows in my datatable 
                int totalRows = dt.Rows.Count;

                // set the first that will be filled with data
                int currentRow = 2;

                // loop through each row.
                for (int i = 0; i < totalRows; i++)
                {
                    // create formulas
                    string controlFormula = string.Format("=IF(AX{0}=0,\" \",\"!!FEIL!!\")", currentRow);
                    string sumPreRoundingFormula = string.Format("=SUM(P{0}:AV{0})", currentRow);
                    string sumRoundedFormula = string.Format("=ROUND(AW{0},2)", currentRow);
                    string vatSales = string.Format("=-(N{0}/1.25)*0.25", currentRow);
                    string salesVATExempt = string.Format("=-(N{0}/1.25)", currentRow);

                    // apply formulas to cells.
                    ws.Cells(string.Format("A{0}", currentRow)).FormulaA1 = controlFormula;
                    ws.Cells(string.Format("AW{0}", currentRow)).FormulaA1 = sumPreRoundingFormula;
                    ws.Cells(string.Format("AX{0}", currentRow)).FormulaA1 = sumRoundedFormula;

                    // add VAT formulas
                    if (ws.Cell(currentRow, 15).Value.Equals("NOK")
                        && (ws.Cell(currentRow, 7).Value.Equals("SHOPIFY"))
                        && (ws.Cell(currentRow, 22).GetValue<decimal>() != 0))
                    {
                        ws.Cells(string.Format("U{0}", currentRow)).FormulaA1 = vatSales;
                        ws.Cells(string.Format("V{0}", currentRow)).FormulaA1 = salesVATExempt;
                    }

                    // increment your counters to apply the same data to the following row
                    currentRow++;
                }

                // set font color for control column
                table.Columns("A").Style.Font.FontColor = XLColor.Red;
                table.Columns("A").Style.Font.Bold = true;

                // set background color for VAT
                var lightGreen = XLColor.FromArgb(0xD8E4BC);
                table.Columns("T:U").Style.Fill.BackgroundColor = lightGreen;

                // set background color for investments, withdrawal and deposits
                var lightBlue = XLColor.FromArgb(0xC5D9F1);
                table.Columns("AS:AV").Style.Fill.BackgroundColor = lightBlue;

                // set background color for control sum
                var lightRed = XLColor.FromArgb(0xE6B8B7);
                table.Columns("AX").Style.Fill.BackgroundColor = lightRed;

                // set column formats
                ws.Range("C2:C9999").Style.NumberFormat.Format = "dd.MM.yyyy";
                ws.Range("E2:E9999").Style.NumberFormat.Format = "####################";

                // Custom formats for numbers in Excel are entered in this format:
                // positive number format;negative number format;zero format;text format
                ws.Range("N2:N9999").Style.NumberFormat.Format = "#,##0.00;[Red]-#,##0.00;";
                ws.Range("N2:N9999").DataType = XLCellValues.Number;

                // set style and format for the decimal range
                var decimalRange = ws.Range("P2:AX9999");
                decimalRange.Style.NumberFormat.Format = "#,##0.00;[Red]-#,##0.00;";
                decimalRange.DataType = XLCellValues.Number;

                // resize
                ws.Columns().AdjustToContents();  // Adjust column width
                ws.Rows().AdjustToContents();     // Adjust row heights

                wb.SaveAs(filePath);
            }
        }

        static List<AccountingItem> ProcessBankAccountStatement(string skandiabankenXLSX, List<string> customerNames)
        {
            var accountingList = new List<AccountingItem>();

            var currentDate = DateTime.Now.Date;
            var currentYear = currentDate.Year;
            var from = new DateTime(currentYear, 1, 1);
            var to = currentDate;

            // prepopulate some lookup lists
            var oberloOrders = Oberlo.GetLatestOberloOrders();
            var aliExpressOrders = AliExpress.GetLatestAliExpressOrders();
            var aliExpressOrderGroups = AliExpress.CombineOrders(aliExpressOrders);

            // run through the bank account transactions
            var skandiabankenBankStatement = Skandiabanken.ReadBankStatement(skandiabankenXLSX);
            var skandiabankenTransactions = skandiabankenBankStatement.Transactions;

            // add incoming balance
            var incomingBalance = new AccountingItem();
            incomingBalance.AccountBank = skandiabankenBankStatement.IncomingBalance;
            incomingBalance.PersonalDeposit = -skandiabankenBankStatement.IncomingBalance;
            incomingBalance.Date = skandiabankenBankStatement.IncomingBalanceDate;
            incomingBalance.Text = skandiabankenBankStatement.IncomingBalanceLabel;
            incomingBalance.AccountingType = "INNGÅENDE SALDO";
            incomingBalance.Type = "Saldo";
            accountingList.Add(incomingBalance);

            // and map each one to the right meta information
            foreach (var skandiabankenTransaction in skandiabankenTransactions)
            {
                // define accounting item
                var accountingItem = new AccountingItem();

                // set date to closer to midnight (sorts better)
                //accountingItem.Date = skandiabankenTransaction.TransactionDate;
                accountingItem.Date = new DateTime(
                    skandiabankenTransaction.TransactionDate.Year,
                    skandiabankenTransaction.TransactionDate.Month,
                    skandiabankenTransaction.TransactionDate.Day,
                    23, 59, 00);

                accountingItem.ArchiveReference = skandiabankenTransaction.ArchiveReference;
                accountingItem.Type = skandiabankenTransaction.Type;

                // extract properties from the transaction text
                skandiabankenTransaction.ExtractAccountingInformation();
                var accountingType = skandiabankenTransaction.AccountingType;
                accountingItem.AccountingType = skandiabankenTransaction.GetAccountingTypeString();

                // 1. If purchase or return from purchase 
                if (skandiabankenTransaction.Type.Equals("Visa") && (
                    accountingType == SkandiabankenTransaction.AccountingTypeEnum.CostOfWebShop ||
                    accountingType == SkandiabankenTransaction.AccountingTypeEnum.CostOfAdvertising ||
                    accountingType == SkandiabankenTransaction.AccountingTypeEnum.CostOfDomain ||
                    accountingType == SkandiabankenTransaction.AccountingTypeEnum.CostOfServer ||
                    accountingType == SkandiabankenTransaction.AccountingTypeEnum.IncomeReturn))
                {

                    Console.WriteLine("{0}", skandiabankenTransaction);
                    accountingItem.Text = string.Format("{0:dd.MM.yyyy} {1} {2} {3} (Kurs: {4})", skandiabankenTransaction.ExternalPurchaseDate, skandiabankenTransaction.ExternalPurchaseVendor, skandiabankenTransaction.ExternalPurchaseAmount, skandiabankenTransaction.ExternalPurchaseCurrency, skandiabankenTransaction.ExternalPurchaseExchangeRate);
                    accountingItem.PurchaseOtherCurrency = skandiabankenTransaction.ExternalPurchaseAmount;
                    accountingItem.OtherCurrency = skandiabankenTransaction.ExternalPurchaseCurrency.ToUpper();
                    accountingItem.AccountBank = skandiabankenTransaction.AccountChange;

                    switch (accountingType)
                    {
                        case SkandiabankenTransaction.AccountingTypeEnum.CostOfWebShop:
                        case SkandiabankenTransaction.AccountingTypeEnum.CostOfDomain:
                        case SkandiabankenTransaction.AccountingTypeEnum.CostOfServer:
                            accountingItem.CostOfData = -skandiabankenTransaction.AccountChange;
                            break;
                        case SkandiabankenTransaction.AccountingTypeEnum.CostOfAdvertising:
                            accountingItem.CostOfAdvertising = -skandiabankenTransaction.AccountChange;
                            break;
                    }
                }

                // 1. If AliExpress purchase
                else if (skandiabankenTransaction.Type.Equals("Visa") &&
                    accountingType == SkandiabankenTransaction.AccountingTypeEnum.CostOfGoods)
                {
                    Console.WriteLine("{0}", skandiabankenTransaction);
                    accountingItem.Text = string.Format("{0:dd.MM.yyyy} {1} {2} {3} (Kurs: {4})", skandiabankenTransaction.ExternalPurchaseDate, skandiabankenTransaction.ExternalPurchaseVendor, skandiabankenTransaction.ExternalPurchaseAmount, skandiabankenTransaction.ExternalPurchaseCurrency, skandiabankenTransaction.ExternalPurchaseExchangeRate);
                    accountingItem.PurchaseOtherCurrency = skandiabankenTransaction.ExternalPurchaseAmount;
                    accountingItem.OtherCurrency = skandiabankenTransaction.ExternalPurchaseCurrency.ToUpper();
                    accountingItem.AccountBank = skandiabankenTransaction.AccountChange;
                    accountingItem.CostForReselling = -skandiabankenTransaction.AccountChange;

                    // lookup in AliExpress purchase list
                    // matching ordertime and orderamount
                    var aliExpressQuery =
                        from order in aliExpressOrderGroups
                        where
                        order.OrderTime.Date == skandiabankenTransaction.ExternalPurchaseDate.Date &&
                        order.OrderAmount == skandiabankenTransaction.ExternalPurchaseAmount
                        orderby order.OrderTime ascending
                        select order;

                    // if the count is more than one, we cannot match easily 
                    if (aliExpressQuery.Count() > 1)
                    {
                        string aliexOrders = String.Join("\n\t", aliExpressQuery.Select(o => o.ToString()));
                        Console.WriteLine("\tERROR: MUST CHOOSE ONE OF MULTIPLE:\n\t{0}", aliexOrders);

                        // flatten the aliexpress order list
                        var aliExpressOrderList = aliExpressQuery.SelectMany(a => a.Children).ToList();

                        // join the aliexpress list and the oberlo list on aliexpress order number
                        var joined = from a in aliExpressOrderList
                                     join b in oberloOrders
                                    on a.OrderId.ToString() equals b.AliOrderNumber
                                     select new { AliExpress = a, Oberlo = b };

                        if (joined.Count() > 0)
                        {
                            // found shopify order numbers
                            Console.WriteLine("\tFOUND SHOPIFY ORDERS:");

                            // join the ordernumbers into a string
                            var orderNumbers = string.Join(", ", joined.Select(c => c.Oberlo).Select(d => d.OrderNumber).Distinct());
                            if (orderNumbers.Equals(""))
                            {
                                accountingItem.ErrorMessage = "Shopify: No orders found";
                                orderNumbers = "NOT FOUND";
                            }
                            else
                            {
                                accountingItem.ErrorMessage = "Shopify: More than one found. Choose one";
                            }
                            Console.WriteLine("\t{0}", orderNumbers);
                            accountingItem.NumPurchase = orderNumbers;
                        }
                        else
                        {
                            // could not find shopify order numbers
                            Console.WriteLine("\tERROR: NO SHOPIFY ORDERS FOUND!");
                            accountingItem.ErrorMessage = "Shopify: No orders found";
                            accountingItem.NumPurchase = "NOT FOUND";
                        }
                    }
                    // one to one match
                    else if (aliExpressQuery.Count() == 1)
                    {
                        Console.WriteLine("\tOK: FOUND SINGLE: {0}", aliExpressQuery.First());

                        // join order ids and make sure they are strings
                        var ids = aliExpressQuery.SelectMany(a => a.Children).Select(b => b.OrderId.ToString()).ToList();

                        // lookup in oberlo to find shopify order number
                        var oberloQuery =
                            from order in oberloOrders
                            where
                            ids.Contains(order.AliOrderNumber)
                            orderby order.CreatedDate ascending
                            select order;

                        // join the ordernumbers into a string
                        var orderNumbers = string.Join(", ", oberloQuery.Select(c => c.OrderNumber).Distinct());
                        if (orderNumbers.Equals(""))
                        {
                            accountingItem.ErrorMessage = "Shopify: No orders found";
                            orderNumbers = "NOT FOUND";
                        }
                        else
                        {
                            // lookup customer name
                            accountingItem.CustomerName = oberloQuery.First().CustomerName;
                        }
                        Console.WriteLine("\t{0}", orderNumbers);
                        accountingItem.NumPurchase = orderNumbers;
                    }
                    else
                    {
                        Console.WriteLine("\tERROR: NO SHOPIFY ORDER FOUND!");
                        accountingItem.ErrorMessage = "Shopify: No orders found";
                        accountingItem.NumPurchase = "NOT FOUND";
                    }
                }

                // 2. Transfer Paypal
                else if (accountingType == SkandiabankenTransaction.AccountingTypeEnum.TransferPaypal)
                {
                    Console.WriteLine("{0}", skandiabankenTransaction);
                    accountingItem.Text = string.Format("{0:dd.MM.yyyy} {1}", skandiabankenTransaction.ExternalPurchaseDate, skandiabankenTransaction.ExternalPurchaseVendor);
                    accountingItem.Gateway = "paypal";

                    accountingItem.AccountPaypal = -skandiabankenTransaction.AccountChange;
                    accountingItem.AccountBank = skandiabankenTransaction.AccountChange;
                }

                // 3. Transfer Stripe
                else if (accountingType == SkandiabankenTransaction.AccountingTypeEnum.TransferStripe)
                {
                    Console.WriteLine("{0}", skandiabankenTransaction);
                    accountingItem.Text = string.Format("{0:dd.MM.yyyy} {1}", skandiabankenTransaction.ExternalPurchaseDate, skandiabankenTransaction.ExternalPurchaseVendor);
                    accountingItem.Gateway = "stripe";

                    accountingItem.AccountStripe = -skandiabankenTransaction.AccountChange;
                    accountingItem.AccountBank = skandiabankenTransaction.AccountChange;
                }

                else if (customerNames.Contains(skandiabankenTransaction.Text))
                {
                    Console.WriteLine("{0}", skandiabankenTransaction);
                    accountingItem.Text = string.Format("{0}", skandiabankenTransaction.Text);
                    accountingItem.Gateway = "vipps";
                    accountingItem.AccountingType = "OVERFØRSEL VIPPS";
                    accountingItem.AccountBank = skandiabankenTransaction.AccountChange;
                    accountingItem.AccountVipps = -skandiabankenTransaction.AccountChange;
                }

                // 4. None of those above
                else
                {
                    Console.WriteLine("{0}", skandiabankenTransaction);
                    accountingItem.Text = string.Format("{0}", skandiabankenTransaction.Text);
                    accountingItem.AccountBank = skandiabankenTransaction.AccountChange;

                    switch (accountingType)
                    {
                        case SkandiabankenTransaction.AccountingTypeEnum.CostOfWebShop:
                        case SkandiabankenTransaction.AccountingTypeEnum.CostOfDomain:
                        case SkandiabankenTransaction.AccountingTypeEnum.CostOfServer:
                            accountingItem.CostOfData = -skandiabankenTransaction.AccountChange;
                            break;
                        case SkandiabankenTransaction.AccountingTypeEnum.CostOfAdvertising:
                            accountingItem.CostOfAdvertising = -skandiabankenTransaction.AccountChange;
                            break;
                        case SkandiabankenTransaction.AccountingTypeEnum.CostOfTryouts:
                            accountingItem.CostOfGoods = -skandiabankenTransaction.AccountChange;
                            break;
                        case SkandiabankenTransaction.AccountingTypeEnum.CostOfBank:
                            accountingItem.CostOfFinance = -skandiabankenTransaction.AccountChange;
                            break;
                        case SkandiabankenTransaction.AccountingTypeEnum.IncomeInterest:
                            accountingItem.IncomeFinance = -skandiabankenTransaction.AccountChange;
                            break;
                        case SkandiabankenTransaction.AccountingTypeEnum.IncomeReturn:
                            accountingItem.CostForReselling = -skandiabankenTransaction.AccountChange;
                            break;
                    }
                }

                accountingList.Add(accountingItem);
            }
            return accountingList;
        }

        static List<AccountingItem> ProcessShopifyStatement(List<string> customerNames)
        {
            var accountingList = new List<AccountingItem>();

            // prepopulate lookup lists
            Console.Out.WriteLine("Prepopulating Lookup Lists ...");

            var stripeTransactions = Stripe.GetLatestStripeTransactions();
            Console.Out.WriteLine("Successfully read Stripe transactions ...");

            var paypalTransactions = Paypal.GetLatestPaypalTransactions();
            Console.Out.WriteLine("Successfully read PayPal transactions ...");

            // get shopify configuration parameters
            string shopifyDomain = ConfigurationManager.AppSettings["ShopifyDomain"];
            string shopifyAPIKey = ConfigurationManager.AppSettings["ShopifyAPIKey"];
            string shopifyAPIPassword = ConfigurationManager.AppSettings["ShopifyAPIPassword"];

            var shopifyOrders = Shopify.ReadShopifyOrders(shopifyDomain, shopifyAPIKey, shopifyAPIPassword);
            Console.Out.WriteLine("Successfully read all Shopify orders ...");

            Console.Out.WriteLine("Processing Shopify orders started ...");
            foreach (var shopifyOrder in shopifyOrders)
            {
                // skip, not paid (pending), cancelled (voided) and fully refunded orders (refunded)
                if (shopifyOrder.FinancialStatus.Equals("refunded")
                    || shopifyOrder.FinancialStatus.Equals("voided")
                    || shopifyOrder.FinancialStatus.Equals("pending")) continue;

                // define accounting item
                var accountingItem = new AccountingItem();
                accountingItem.Date = shopifyOrder.CreatedAt;
                accountingItem.ArchiveReference = shopifyOrder.Id;
                accountingItem.Type = string.Format("{0} {1}", shopifyOrder.FinancialStatus, shopifyOrder.FulfillmentStatus);
                accountingItem.AccountingType = "SHOPIFY";
                accountingItem.Text = string.Format("SALG {0} {1}", shopifyOrder.CustomerName, shopifyOrder.PaymentId);
                accountingItem.CustomerName = shopifyOrder.CustomerName;

                // add to customer name list
                customerNames.Add(accountingItem.CustomerName);

                if (shopifyOrder.Gateway != null)
                {
                    accountingItem.Gateway = shopifyOrder.Gateway.ToLower();
                }
                accountingItem.NumSale = shopifyOrder.Name;

                var startDate = shopifyOrder.ProcessedAt.AddDays(-1);
                var endDate = shopifyOrder.ProcessedAt.AddDays(1);

                switch (accountingItem.Gateway)
                {
                    case "vipps":
                        accountingItem.PurchaseOtherCurrency = shopifyOrder.TotalPrice;
                        accountingItem.OtherCurrency = "NOK";

                        //accountingItem.FeesVipps = fee;
                        accountingItem.AccountVipps = shopifyOrder.TotalPrice;

                        break;
                    case "stripe":

                        accountingItem.PurchaseOtherCurrency = shopifyOrder.TotalPrice;
                        accountingItem.OtherCurrency = "NOK";

                        // lookup the stripe transaction
                        var stripeQuery =
                        from transaction in stripeTransactions
                        where
                        transaction.Paid &&
                        transaction.CustomerEmail.Equals(shopifyOrder.CustomerEmail) &&
                        transaction.Amount == shopifyOrder.TotalPrice &&
                         (transaction.Created.Date >= startDate.Date && transaction.Created.Date <= endDate.Date)
                        orderby transaction.Created ascending
                        select transaction;

                        if (stripeQuery.Count() > 1)
                        {
                            // more than one ?!
                            Console.Out.WriteLine("ERROR: FOUND MORE THAN ONE MATCHING STRIPE TRANSACTION!");
                            accountingItem.ErrorMessage = "Stripe: More than one found, choose one";
                        }
                        else if (stripeQuery.Count() > 0)
                        {
                            // one match
                            var stripeTransaction = stripeQuery.First();
                            decimal amount = stripeTransaction.Amount;
                            decimal net = stripeTransaction.Net;
                            decimal fee = stripeTransaction.Fee;

                            accountingItem.FeesStripe = fee;
                            accountingItem.AccountStripe = net;
                        }
                        else
                        {
                            Console.Out.WriteLine("ERROR: NO STRIPE TRANSACTIONS FOUND!");
                            accountingItem.ErrorMessage = "Stripe: No transactions found";
                        }

                        break;
                    case "paypal":

                        accountingItem.PurchaseOtherCurrency = shopifyOrder.TotalPrice;
                        accountingItem.OtherCurrency = "NOK";

                        // lookup the paypal transaction
                        var paypalQuery =
                        from transaction in paypalTransactions
                        let grossAmount = transaction.GrossAmount
                        let timestamp = transaction.Timestamp
                        where
                        transaction.Status.Equals("Completed")
                        //&& (null != transaction.Payer && transaction.Payer.Equals(shopifyOrder.CustomerEmail))
                        && (
                        (null != transaction.PayerDisplayName && transaction.PayerDisplayName.Equals(shopifyOrder.CustomerName, StringComparison.InvariantCultureIgnoreCase))
                        ||
                        (null != transaction.Payer && transaction.Payer.Equals(shopifyOrder.CustomerEmail, StringComparison.InvariantCultureIgnoreCase))
                        )
                        && (grossAmount == shopifyOrder.TotalPrice)
                        && (timestamp.Date >= startDate.Date && timestamp.Date <= endDate.Date)
                        //&& (timestamp.Date == shopifyOrder.ProcessedAt.Date)
                        orderby timestamp ascending
                        select transaction;

                        if (paypalQuery.Count() > 1)
                        {
                            // more than one ?!
                            Console.Out.WriteLine("ERROR: FOUND MORE THAN ONE PAYPAL TRANSACTION!");
                            accountingItem.ErrorMessage = "Paypal: More than one found, choose one";
                        }
                        else if (paypalQuery.Count() > 0)
                        {
                            // one match
                            var paypalTransaction = paypalQuery.First();
                            decimal amount = paypalTransaction.GrossAmount;
                            decimal net = paypalTransaction.NetAmount;
                            decimal fee = paypalTransaction.FeeAmount;

                            accountingItem.FeesPaypal = -fee;
                            accountingItem.AccountPaypal = net;
                        }
                        else
                        {
                            Console.Out.WriteLine("ERROR: NO PAYPAL TRANSACTIONS FOUND!");
                            accountingItem.ErrorMessage = "Paypal: No transactions found";
                        }

                        break;
                }

                // fix VAT
                if (shopifyOrder.TotalTax != 0)
                {
                    accountingItem.SalesVAT = -(shopifyOrder.TotalPrice / (decimal)1.25);
                    accountingItem.VATSales = accountingItem.SalesVAT * (decimal)0.25;
                }
                else
                {
                    accountingItem.SalesVATExempt = -shopifyOrder.TotalPrice;
                }

                // check if free gift
                if (shopifyOrder.TotalPrice == 0)
                {
                    accountingItem.AccountingType += " FREE";
                    accountingItem.Gateway = "none";
                }

                accountingList.Add(accountingItem);
            }

            return accountingList;
        }
    }
}
