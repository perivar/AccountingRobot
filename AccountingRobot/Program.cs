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

            //var fileName = @"C:\Users\pnerseth\Amazon Drive\Documents\Private\wazalo\regnskap\Accounting Fixed 2017-12-22.xlsx";
            //UpdateExcelFile(fileName, accountingItems);

            Console.ReadLine();
        }

        #region Excel Methods
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
                table.Theme = XLTableTheme.TableStyleLight16;

                // turn on table total rows and set the functions for each of the relevant columns
                SetExcelTableTotalsRowFunction(table);

                if (table != null)
                {
                    foreach (var row in table.DataRange.Rows())
                    {
                        SetExcelRowFormulas(row);
                        SetExcelRowStyles(row);
                    }
                }

                // resize
                ws.Columns().AdjustToContents();  // Adjust column width
                ws.Rows().AdjustToContents();     // Adjust row heights

                wb.SaveAs(filePath);
            }
        }

        static void UpdateExcelFile(string filePath, List<AccountingItem> newAccountingItems)
        {
            // go through each row and check if it has already been "fixed".
            // i.e. the Number columns is no longer 0

            XLWorkbook wb = new XLWorkbook(filePath);
            IXLWorksheet ws = wb.Worksheet("Accounting");

            IXLTables tables = ws.Tables;
            IXLTable table = tables.FirstOrDefault();

            var oldAccountingSpreadsheet = new List<AccountingItem>();
            if (table != null)
            {
                foreach (var row in table.DataRange.Rows())
                {
                    var accountingItem = new AccountingItem();

                    accountingItem.Date = ExcelUtils.GetExcelField<DateTime>(row, "Date");
                    accountingItem.Number = ExcelUtils.GetExcelField<int>(row, "Number");
                    accountingItem.ArchiveReference = ExcelUtils.GetExcelField<long>(row, "ArchiveReference");
                    accountingItem.Type = ExcelUtils.GetExcelField<string>(row, "Type");
                    accountingItem.AccountingType = ExcelUtils.GetExcelField<string>(row, "AccountingType");
                    accountingItem.Text = ExcelUtils.GetExcelField<string>(row, "Text");
                    accountingItem.CustomerName = ExcelUtils.GetExcelField<string>(row, "CustomerName");
                    accountingItem.ErrorMessage = ExcelUtils.GetExcelField<string>(row, "ErrorMessage");
                    accountingItem.Gateway = ExcelUtils.GetExcelField<string>(row, "Gateway");
                    accountingItem.NumSale = ExcelUtils.GetExcelField<string>(row, "NumSale");
                    accountingItem.NumPurchase = ExcelUtils.GetExcelField<string>(row, "NumPurchase");
                    accountingItem.PurchaseOtherCurrency = ExcelUtils.GetExcelField<decimal>(row, "PurchaseOtherCurrency");
                    accountingItem.OtherCurrency = ExcelUtils.GetExcelField<string>(row, "OtherCurrency");

                    accountingItem.AccountPaypal = ExcelUtils.GetExcelField<decimal>(row, "AccountPaypal");	// 1910
                    accountingItem.AccountStripe = ExcelUtils.GetExcelField<decimal>(row, "AccountStripe");	// 1915
                    accountingItem.AccountVipps = ExcelUtils.GetExcelField<decimal>(row, "AccountVipps");	// 1918
                    accountingItem.AccountBank = ExcelUtils.GetExcelField<decimal>(row, "AccountBank");	// 1920

                    accountingItem.VATPurchase = ExcelUtils.GetExcelField<decimal>(row, "VATPurchase");
                    accountingItem.VATSales = ExcelUtils.GetExcelField<decimal>(row, "VATSales");

                    accountingItem.SalesVAT = ExcelUtils.GetExcelField<decimal>(row, "SalesVAT");	// 3000
                    accountingItem.SalesVATExempt = ExcelUtils.GetExcelField<decimal>(row, "SalesVATExempt");	// 3100

                    accountingItem.CostOfGoods = ExcelUtils.GetExcelField<decimal>(row, "CostOfGoods");	// 4005
                    accountingItem.CostForReselling = ExcelUtils.GetExcelField<decimal>(row, "CostForReselling");	// 4300
                    accountingItem.CostForSalary = ExcelUtils.GetExcelField<decimal>(row, "CostForSalary");	// 5000
                    accountingItem.CostForSalaryTax = ExcelUtils.GetExcelField<decimal>(row, "CostForSalaryTax");	// 5400
                    accountingItem.CostForDepreciation = ExcelUtils.GetExcelField<decimal>(row, "CostForDepreciation");	// 6000
                    accountingItem.CostForShipping = ExcelUtils.GetExcelField<decimal>(row, "CostForShipping");	// 6100
                    accountingItem.CostForElectricity = ExcelUtils.GetExcelField<decimal>(row, "CostForElectricity");	// 6340 
                    accountingItem.CostForToolsInventory = ExcelUtils.GetExcelField<decimal>(row, "CostForToolsInventory");	// 6500
                    accountingItem.CostForMaintenance = ExcelUtils.GetExcelField<decimal>(row, "CostForMaintenance");	// 6695
                    accountingItem.CostForFacilities = ExcelUtils.GetExcelField<decimal>(row, "CostForFacilities");	// 6800 

                    accountingItem.CostOfData = ExcelUtils.GetExcelField<decimal>(row, "CostOfData");	// 6810 
                    accountingItem.CostOfPhoneInternet = ExcelUtils.GetExcelField<decimal>(row, "CostOfPhoneInternet");	// 6900
                    accountingItem.CostForTravelAndAllowance = ExcelUtils.GetExcelField<decimal>(row, "CostForTravelAndAllowance");	// 7140
                    accountingItem.CostOfAdvertising = ExcelUtils.GetExcelField<decimal>(row, "CostOfAdvertising");	// 7330
                    accountingItem.CostOfOther = ExcelUtils.GetExcelField<decimal>(row, "CostOfOther");	// 7700

                    accountingItem.FeesBank = ExcelUtils.GetExcelField<decimal>(row, "FeesBank");	// 7770
                    accountingItem.FeesPaypal = ExcelUtils.GetExcelField<decimal>(row, "FeesPaypal");	// 7780
                    accountingItem.FeesStripe = ExcelUtils.GetExcelField<decimal>(row, "FeesStripe");	// 7785 

                    accountingItem.CostForEstablishment = ExcelUtils.GetExcelField<decimal>(row, "CostForEstablishment");	// 7790

                    accountingItem.IncomeFinance = ExcelUtils.GetExcelField<decimal>(row, "IncomeFinance");	// 8099
                    accountingItem.CostOfFinance = ExcelUtils.GetExcelField<decimal>(row, "CostOfFinance");	// 8199

                    accountingItem.Investments = ExcelUtils.GetExcelField<decimal>(row, "Investments");	// 1200
                    accountingItem.AccountsReceivable = ExcelUtils.GetExcelField<decimal>(row, "AccountsReceivable");	// 1500
                    accountingItem.PersonalWithdrawal = ExcelUtils.GetExcelField<decimal>(row, "PersonalWithdrawal");
                    accountingItem.PersonalDeposit = ExcelUtils.GetExcelField<decimal>(row, "PersonalDeposit");

                    oldAccountingSpreadsheet.Add(accountingItem);
                }

                // turn off totals row before adding more rows
                table.ShowTotalsRow = false;

                // add more rows
                var firstNotSecond = newAccountingItems.Except(oldAccountingSpreadsheet).ToList();
                var secondNotFirst = oldAccountingSpreadsheet.Except(newAccountingItems).ToList();
                var newRows = table.InsertRowsBelow(firstNotSecond.Count(), true);

                var counter = 0;
                foreach (var newRow in newRows)
                {
                    newRow.Cell(1).Value = "";
                    newRow.Cell(2).Value = firstNotSecond[counter].Periode;
                    newRow.Cell(3).Value = firstNotSecond[counter].Date;
                    newRow.Cell(4).Value = firstNotSecond[counter].Number;
                    newRow.Cell(5).Value = firstNotSecond[counter].ArchiveReference;
                    newRow.Cell(6).Value = firstNotSecond[counter].Type;
                    newRow.Cell(7).Value = firstNotSecond[counter].AccountingType;
                    newRow.Cell(8).Value = firstNotSecond[counter].Text;
                    newRow.Cell(9).Value = firstNotSecond[counter].CustomerName;
                    newRow.Cell(10).Value = firstNotSecond[counter].ErrorMessage;
                    newRow.Cell(11).Value = firstNotSecond[counter].Gateway;
                    newRow.Cell(12).Value = firstNotSecond[counter].NumSale;
                    newRow.Cell(13).Value = firstNotSecond[counter].NumPurchase;
                    newRow.Cell(14).Value = firstNotSecond[counter].PurchaseOtherCurrency;
                    newRow.Cell(15).Value = firstNotSecond[counter].OtherCurrency;

                    newRow.Cell(16).Value = firstNotSecond[counter].AccountPaypal;               // 1910
                    newRow.Cell(17).Value = firstNotSecond[counter].AccountStripe;               // 1915
                    newRow.Cell(18).Value = firstNotSecond[counter].AccountVipps;                // 1918
                    newRow.Cell(19).Value = firstNotSecond[counter].AccountBank;                 // 1920

                    newRow.Cell(20).Value = firstNotSecond[counter].VATPurchase;
                    newRow.Cell(21).Value = firstNotSecond[counter].VATSales;

                    newRow.Cell(22).Value = firstNotSecond[counter].SalesVAT;                    // 3000
                    newRow.Cell(23).Value = firstNotSecond[counter].SalesVATExempt;              // 3100

                    newRow.Cell(24).Value = firstNotSecond[counter].CostOfGoods;                 // 4005
                    newRow.Cell(25).Value = firstNotSecond[counter].CostForReselling;            // 4300
                    newRow.Cell(26).Value = firstNotSecond[counter].CostForSalary;               // 5000
                    newRow.Cell(27).Value = firstNotSecond[counter].CostForSalaryTax;            // 5400
                    newRow.Cell(28).Value = firstNotSecond[counter].CostForDepreciation;         // 6000
                    newRow.Cell(29).Value = firstNotSecond[counter].CostForShipping;             // 6100
                    newRow.Cell(30).Value = firstNotSecond[counter].CostForElectricity;          // 6340 
                    newRow.Cell(31).Value = firstNotSecond[counter].CostForToolsInventory;       // 6500
                    newRow.Cell(32).Value = firstNotSecond[counter].CostForMaintenance;          // 6695
                    newRow.Cell(33).Value = firstNotSecond[counter].CostForFacilities;           // 6800 

                    newRow.Cell(34).Value = firstNotSecond[counter].CostOfData;                  // 6810 
                    newRow.Cell(35).Value = firstNotSecond[counter].CostOfPhoneInternet;         // 6900
                    newRow.Cell(36).Value = firstNotSecond[counter].CostForTravelAndAllowance;   // 7140
                    newRow.Cell(37).Value = firstNotSecond[counter].CostOfAdvertising;           // 7330
                    newRow.Cell(38).Value = firstNotSecond[counter].CostOfOther;                 // 7700

                    newRow.Cell(39).Value = firstNotSecond[counter].FeesBank;                    // 7770
                    newRow.Cell(40).Value = firstNotSecond[counter].FeesPaypal;                  // 7780
                    newRow.Cell(41).Value = firstNotSecond[counter].FeesStripe;                  // 7785 

                    newRow.Cell(42).Value = firstNotSecond[counter].CostForEstablishment;        // 7790

                    newRow.Cell(43).Value = firstNotSecond[counter].IncomeFinance;               // 8099
                    newRow.Cell(44).Value = firstNotSecond[counter].CostOfFinance;               // 8199

                    newRow.Cell(45).Value = firstNotSecond[counter].Investments;                 // 1200
                    newRow.Cell(46).Value = firstNotSecond[counter].AccountsReceivable;          // 1500
                    newRow.Cell(47).Value = firstNotSecond[counter].PersonalWithdrawal;
                    newRow.Cell(48).Value = firstNotSecond[counter].PersonalDeposit;

                    SetExcelRowFormulas(newRow);
                    SetExcelRowStyles(newRow);

                    counter++;
                }

                // turn on table total rows and set the functions for each of the relevant columns
                SetExcelTableTotalsRowFunction(table);
            }

            // resize
            ws.Columns().AdjustToContents();  // Adjust column width
            ws.Rows().AdjustToContents();     // Adjust row heights

            wb.SaveAs(@"test.xlsx");
        }

        static void SetExcelRowFormulas(IXLRangeRow row)
        {
            int currentRow = row.RowNumber();

            // create formulas
            string controlFormula = string.Format("=IF(AX{0}=0,\" \",\"!!FEIL!!\")", currentRow);
            string sumPreRoundingFormula = string.Format("=SUM(P{0}:AV{0})", currentRow);
            string sumRoundedFormula = string.Format("=ROUND(AW{0},2)", currentRow);
            string vatSales = string.Format("=-(N{0}/1.25)*0.25", currentRow);
            string salesVATExempt = string.Format("=-(N{0}/1.25)", currentRow);

            // apply formulas to cells.
            row.Cell("A").FormulaA1 = controlFormula;
            row.Cell("AW").FormulaA1 = sumPreRoundingFormula;
            row.Cell("AX").FormulaA1 = sumRoundedFormula;

            // add VAT formulas
            if (row.Cell("O").Value.Equals("NOK")
                && (row.Cell("G").Value.Equals("SHOPIFY"))
                && (row.Cell("V").GetValue<decimal>() != 0))
            {
                row.Cell("U").FormulaA1 = vatSales;
                row.Cell("V").FormulaA1 = salesVATExempt;
            }
        }

        static void SetExcelRowStyles(IXLRangeRow row)
        {
            int currentRow = row.RowNumber();

            // set font color for control column
            row.Cell("A").Style.Font.FontColor = XLColor.Red;
            row.Cell("A").Style.Font.Bold = true;

            // set background color for VAT
            var lightGreen = XLColor.FromArgb(0xD8E4BC);
            var lighterGreen = XLColor.FromArgb(0xEBF1DE);
            var green = currentRow % 2 == 0 ? lightGreen : lighterGreen;
            row.Cells("T","U").Style.Fill.BackgroundColor = green;

            // set background color for investments, withdrawal and deposits
            var lightBlue = XLColor.FromArgb(0xC5D9F1);
            var lighterBlue = XLColor.FromArgb(0xEAF1FA); 
             var blue = currentRow % 2 == 0 ? lightBlue : lighterBlue;
            row.Cells("AS","AV").Style.Fill.BackgroundColor = blue;

            // set background color for control sum
            var lightRed = XLColor.FromArgb(0xE6B8B7);
            var lighterRed = XLColor.FromArgb(0xF2DCDB);
            var red = currentRow % 2 == 0 ? lightRed : lighterRed;
            row.Cell("AX").Style.Fill.BackgroundColor = red;

            // set column formats
            row.Cell("C").Style.NumberFormat.Format = "dd.MM.yyyy";
            row.Cell("E").Style.NumberFormat.Format = "####################";

            // Custom formats for numbers in Excel are entered in this format:
            // positive number format;negative number format;zero format;text format
            row.Cell("N").Style.NumberFormat.Format = "#,##0.00;[Red]-#,##0.00;";
            row.Cell("N").DataType = XLCellValues.Number;

            // set style and format for the decimal range
            row.Cells("P","AX").Style.NumberFormat.Format = "#,##0.00;[Red]-#,##0.00;";
            row.Cells("P","AX").DataType = XLCellValues.Number;
        }

        static void SetExcelTableTotalsRowFunction(IXLTable table)
        {
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

        }
        #endregion

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

            var usedOrderNumbers = new HashSet<string>();

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

                    FindAliExpressOrderNumber(usedOrderNumbers, aliExpressOrderGroups, oberloOrders, skandiabankenTransaction, accountingItem);
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
                            Console.Out.WriteLine("ERROR: NO STRIPE TRANSACTIONS FOR {0:C} FOUND FOR {1} {2} BETWEEN {3:dd.MM.yyyy} and {4:dd.MM.yyyy}!", shopifyOrder.TotalPrice, shopifyOrder.Name, shopifyOrder.CustomerName, startDate, endDate);
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
                            Console.Out.WriteLine("ERROR: NO PAYPAL TRANSACTIONS FOR {0:C} FOUND FOR {1} {2} BETWEEN {3:dd.MM.yyyy} and {4:dd.MM.yyyy}!", shopifyOrder.TotalPrice, shopifyOrder.Name, shopifyOrder.CustomerName, startDate, endDate);
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

        #region AliExpress Methods
        static void FindAliExpressOrderNumber(HashSet<string> usedOrderNumbers, List<AliExpressOrderGroup> aliExpressOrderGroups, List<OberloOrder> oberloOrders, SkandiabankenTransaction skandiabankenTransaction, AccountingItem accountingItem)
        {
            // set start and stop date
            var startDate = skandiabankenTransaction.ExternalPurchaseDate.AddDays(-4);
            var endDate = skandiabankenTransaction.ExternalPurchaseDate;

            // lookup in AliExpress purchase list
            // matching ordertime and orderamount
            var aliExpressQuery =
                from order in aliExpressOrderGroups
                where
                (order.OrderTime.Date >= startDate.Date && order.OrderTime.Date <= endDate.Date) &&
                order.OrderAmount == skandiabankenTransaction.ExternalPurchaseAmount
                orderby order.OrderTime ascending
                select order;

            // if the count is more than one, we cannot match easily 
            if (aliExpressQuery.Count() > 1)
            {
                // first check if one of the found orders was ordered on the given purchase date
                var aliExpressQueryExactDate =
                from order in aliExpressQuery
                where
                order.OrderTime.Date == skandiabankenTransaction.ExternalPurchaseDate.Date
                orderby order.OrderTime ascending
                select order;

                // if the count is only one, we have a single match
                if (aliExpressQueryExactDate.Count() == 1)
                {
                    ProcessAliExpressMatch(usedOrderNumbers, aliExpressQueryExactDate, oberloOrders, accountingItem);
                    return;
                }
                // use the original query and present the results
                else
                {
                    ProcessAliExpressMatch(usedOrderNumbers, aliExpressQuery, oberloOrders, accountingItem);
                }
            }
            // if the count is only one, we have a single match
            else if (aliExpressQuery.Count() == 1)
            {
                ProcessAliExpressMatch(usedOrderNumbers, aliExpressQuery, oberloOrders, accountingItem);
            }
            // no orders found
            else
            {
                // could not find shopify order numbers
                Console.WriteLine("\tERROR: NO SHOPIFY ORDERS FOUND!");
                accountingItem.ErrorMessage = "Shopify: No orders found";
                accountingItem.NumPurchase = "NOT FOUND";
            }
        }

        static void ProcessAliExpressMatch(HashSet<string> usedOrderNumbers, IOrderedEnumerable<AliExpressOrderGroup> aliExpressQuery, List<OberloOrder> oberloOrders, AccountingItem accountingItem)
        {
            // flatten the aliexpress order list
            var aliExpressOrderList = aliExpressQuery.SelectMany(a => a.Children).ToList();

            // join the aliexpress list and the oberlo list on aliexpress order number
            var joined = from a in aliExpressOrderList
                         join b in oberloOrders
                        on a.OrderId.ToString() equals b.AliOrderNumber
                         select new { AliExpress = a, Oberlo = b };

            if (joined.Count() > 0)
            {
                Console.WriteLine("\tSHOPIFY ORDERS FOUND ...");

                string orderNumber = "NONE FOUND";
                foreach (var item in joined.Reverse())
                {
                    orderNumber = item.Oberlo.OrderNumber;
                    if (!usedOrderNumbers.Contains(orderNumber))
                    {
                        usedOrderNumbers.Add(orderNumber);
                        accountingItem.NumPurchase = orderNumber;
                        accountingItem.CustomerName = item.Oberlo.CustomerName;
                        Console.WriteLine("\tSELECTED: {0} {1}", accountingItem.NumPurchase, accountingItem.CustomerName);
                        break;
                    }
                }
            }

            // could not find shopify order numbers
            else
            {
                Console.WriteLine("\tERROR: NO OBERLO ORDERS FOUND!");
                accountingItem.ErrorMessage = "Oberlo: No orders found";
                accountingItem.NumPurchase = "NOT FOUND";
            }
        }
        #endregion
    }
}
