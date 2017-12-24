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
            //var now = DateTime.Now;
            //var fileName = string.Format("Accounting {0:yyyy-MM-dd}.xlsx", now);
            //ExportToExcel(fileName, accountingItems);

            var fileName = @"C:\Users\pnerseth\Amazon Drive\Documents\Private\wazalo\regnskap\Accounting 2017-12-24.xlsx";
            UpdateExcelFile(fileName, accountingItems);

            Console.ReadLine();
        }

        #region Excel Methods
        static void ExportToExcel(string filePath, List<AccountingItem> accountingItems)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Kontroll", typeof(String));

            dt.Columns.Add("Periode", typeof(int));
            dt.Columns.Add("Dato", typeof(DateTime));
            dt.Columns.Add("Bilagsnr.", typeof(int));
            dt.Columns.Add("Arkivreferanse", typeof(long));
            dt.Columns.Add("Type", typeof(string));
            dt.Columns.Add("Regnskapstype", typeof(string));
            dt.Columns.Add("Tekst", typeof(string));
            dt.Columns.Add("Kundenavn", typeof(string));
            dt.Columns.Add("Feilmelding", typeof(string));
            dt.Columns.Add("Gateway", typeof(string));
            dt.Columns.Add("Num Salg", typeof(string));
            dt.Columns.Add("Num Kjøp", typeof(string));
            dt.Columns.Add("Kjøp annen valuta", typeof(decimal));
            dt.Columns.Add("Annen valuta", typeof(string));

            dt.Columns.Add("Paypal", typeof(decimal));           // 1910
            dt.Columns.Add("Stripe", typeof(decimal));           // 1915
            dt.Columns.Add("Vipps", typeof(decimal));            // 1918
            dt.Columns.Add("Bank", typeof(decimal));             // 1920

            dt.Columns.Add("MVA Kjøp", typeof(decimal));
            dt.Columns.Add("MVA Salg", typeof(decimal));

            dt.Columns.Add("Salg mva-pliktig", typeof(decimal));                // 3000
            dt.Columns.Add("Salg avgiftsfritt", typeof(decimal));          // 3100

            dt.Columns.Add("Varekostnad", typeof(decimal));             // 4005
            dt.Columns.Add("Forbruk for videresalg", typeof(decimal));        // 4300
            dt.Columns.Add("Lønn", typeof(decimal));           // 5000
            dt.Columns.Add("Arb.giver avgift", typeof(decimal));        // 5400
            dt.Columns.Add("Avskrivninger", typeof(decimal));     // 6000
            dt.Columns.Add("Frakt", typeof(decimal));         // 6100
            dt.Columns.Add("Strøm", typeof(decimal));      // 6340 
            dt.Columns.Add("Verktøy inventar", typeof(decimal));   // 6500
            dt.Columns.Add("Vedlikehold", typeof(decimal));      // 6695
            dt.Columns.Add("Kontorkostnader", typeof(decimal));       // 6800 

            dt.Columns.Add("Datakostnader", typeof(decimal));              // 6810 
            dt.Columns.Add("Telefon Internett", typeof(decimal));     // 6900
            dt.Columns.Add("Reise og Diett", typeof(decimal));// 7140
            dt.Columns.Add("Reklamekostnader", typeof(decimal));       // 7330
            dt.Columns.Add("Diverse annet", typeof(decimal));             // 7700

            dt.Columns.Add("Gebyrer Bank", typeof(decimal));                // 7770
            dt.Columns.Add("Gebyrer Paypal", typeof(decimal));              // 7780
            dt.Columns.Add("Gebyrer Stripe", typeof(decimal));              // 7785 

            dt.Columns.Add("Etableringskostnader", typeof(decimal));    // 7790

            dt.Columns.Add("Finansinntekter", typeof(decimal));           // 8099
            dt.Columns.Add("Finanskostnader", typeof(decimal));           // 8199

            dt.Columns.Add("Investeringer", typeof(decimal));             // 1200
            dt.Columns.Add("Kundefordringer", typeof(decimal));      // 1500
            dt.Columns.Add("Privat uttak", typeof(decimal));
            dt.Columns.Add("Privat innskudd", typeof(decimal));

            dt.Columns.Add("Sum før avrunding", typeof(decimal));
            dt.Columns.Add("Sum", typeof(decimal));

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
                var ws = wb.Worksheets.Add(dt, "Bilagsjournal");
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
                    accountingItem.ArchiveReference = ExcelUtils.GetExcelField<long>(row, "Arkivreferanse");
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

                    accountingItem.AccountPaypal = ExcelUtils.GetExcelField<decimal>(row, "Paypal");	// 1910
                    accountingItem.AccountStripe = ExcelUtils.GetExcelField<decimal>(row, "Stripe");	// 1915
                    accountingItem.AccountVipps = ExcelUtils.GetExcelField<decimal>(row, "Vipps");	// 1918
                    accountingItem.AccountBank = ExcelUtils.GetExcelField<decimal>(row, "Bank");	// 1920

                    accountingItem.VATPurchase = ExcelUtils.GetExcelField<decimal>(row, "MVA Kjøp");
                    accountingItem.VATSales = ExcelUtils.GetExcelField<decimal>(row, "MVA Salg");

                    accountingItem.SalesVAT = ExcelUtils.GetExcelField<decimal>(row, "Salg mva-pliktig");	// 3000
                    accountingItem.SalesVATExempt = ExcelUtils.GetExcelField<decimal>(row, "Salg avgiftsfritt");	// 3100

                    accountingItem.CostOfGoods = ExcelUtils.GetExcelField<decimal>(row, "Varekostnad");	// 4005
                    accountingItem.CostForReselling = ExcelUtils.GetExcelField<decimal>(row, "Forbruk for videresalg");	// 4300
                    accountingItem.CostForSalary = ExcelUtils.GetExcelField<decimal>(row, "Lønn");	// 5000
                    accountingItem.CostForSalaryTax = ExcelUtils.GetExcelField<decimal>(row, "Arb.giver avgift");	// 5400
                    accountingItem.CostForDepreciation = ExcelUtils.GetExcelField<decimal>(row, "Avskrivninger");	// 6000
                    accountingItem.CostForShipping = ExcelUtils.GetExcelField<decimal>(row, "Frakt");	// 6100
                    accountingItem.CostForElectricity = ExcelUtils.GetExcelField<decimal>(row, "Strøm");	// 6340 
                    accountingItem.CostForToolsInventory = ExcelUtils.GetExcelField<decimal>(row, "Verktøy inventar");	// 6500
                    accountingItem.CostForMaintenance = ExcelUtils.GetExcelField<decimal>(row, "Vedlikehold");	// 6695
                    accountingItem.CostForFacilities = ExcelUtils.GetExcelField<decimal>(row, "Kontorkostnader");	// 6800 

                    accountingItem.CostOfData = ExcelUtils.GetExcelField<decimal>(row, "Datakostnader");	// 6810 
                    accountingItem.CostOfPhoneInternet = ExcelUtils.GetExcelField<decimal>(row, "Telefon Internett");	// 6900
                    accountingItem.CostForTravelAndAllowance = ExcelUtils.GetExcelField<decimal>(row, "Reise og Diett");	// 7140
                    accountingItem.CostOfAdvertising = ExcelUtils.GetExcelField<decimal>(row, "Reklamekostnader");	// 7330
                    accountingItem.CostOfOther = ExcelUtils.GetExcelField<decimal>(row, "Diverse annet");	// 7700

                    accountingItem.FeesBank = ExcelUtils.GetExcelField<decimal>(row, "Gebyrer Bank");	// 7770
                    accountingItem.FeesPaypal = ExcelUtils.GetExcelField<decimal>(row, "Gebyrer Paypal");	// 7780
                    accountingItem.FeesStripe = ExcelUtils.GetExcelField<decimal>(row, "Gebyrer Stripe");	// 7785 

                    accountingItem.CostForEstablishment = ExcelUtils.GetExcelField<decimal>(row, "Etableringskostnader");	// 7790

                    accountingItem.IncomeFinance = ExcelUtils.GetExcelField<decimal>(row, "Finansinntekter");	// 8099
                    accountingItem.CostOfFinance = ExcelUtils.GetExcelField<decimal>(row, "Finanskostnader");	// 8199

                    accountingItem.Investments = ExcelUtils.GetExcelField<decimal>(row, "Investeringer");	// 1200
                    accountingItem.AccountsReceivable = ExcelUtils.GetExcelField<decimal>(row, "Kundefordringer");	// 1500
                    accountingItem.PersonalWithdrawal = ExcelUtils.GetExcelField<decimal>(row, "Privat uttak");
                    accountingItem.PersonalDeposit = ExcelUtils.GetExcelField<decimal>(row, "Privat innskudd");

                    existingAccountingItems.Add(row, accountingItem);
                }

                // reduce the old Accounting Spreadsheet and remove the entries that doesn't have a number
                var existingAccountingItemsToDelete =
                    (from row in existingAccountingItems
                     where
                     row.Value.Number == 0
                     orderby row.Value.Number ascending
                     select row);

                // identify elements from the new accounting items list that does not exist in the existing spreadsheet
                var existingAccountingItemsToKeep = existingAccountingItems.Except(existingAccountingItemsToDelete);
                var newAccountingElements = newAccountingItems.Except(existingAccountingItemsToKeep.Select(o => o.Value)).ToList();

                // delete rows from table
                foreach (var deleteRow in existingAccountingItemsToDelete) {
                    deleteRow.Key.Delete(XLShiftDeletedCells.ShiftCellsUp);
                }

                // insert new rows below the existing table
                // turn off totals row before adding more rows
                table.ShowTotalsRow = false;
                var newRows = table.InsertRowsBelow(newAccountingElements.Count(), true);

                var counter = 0;
                foreach (var newRow in newRows)
                {
                    newRow.Cell(1).Value = "";
                    newRow.Cell(2).Value = newAccountingElements[counter].Periode;
                    newRow.Cell(3).Value = newAccountingElements[counter].Date;
                    newRow.Cell(4).Value = newAccountingElements[counter].Number;
                    newRow.Cell(5).Value = newAccountingElements[counter].ArchiveReference;
                    newRow.Cell(6).Value = newAccountingElements[counter].Type;
                    newRow.Cell(7).Value = newAccountingElements[counter].AccountingType;
                    newRow.Cell(8).Value = newAccountingElements[counter].Text;
                    newRow.Cell(9).Value = newAccountingElements[counter].CustomerName;
                    newRow.Cell(10).Value = newAccountingElements[counter].ErrorMessage;
                    newRow.Cell(11).Value = newAccountingElements[counter].Gateway;
                    newRow.Cell(12).Value = newAccountingElements[counter].NumSale;
                    newRow.Cell(13).Value = newAccountingElements[counter].NumPurchase;
                    newRow.Cell(14).Value = newAccountingElements[counter].PurchaseOtherCurrency;
                    newRow.Cell(15).Value = newAccountingElements[counter].OtherCurrency;

                    newRow.Cell(16).Value = newAccountingElements[counter].AccountPaypal;               // 1910
                    newRow.Cell(17).Value = newAccountingElements[counter].AccountStripe;               // 1915
                    newRow.Cell(18).Value = newAccountingElements[counter].AccountVipps;                // 1918
                    newRow.Cell(19).Value = newAccountingElements[counter].AccountBank;                 // 1920

                    newRow.Cell(20).Value = newAccountingElements[counter].VATPurchase;
                    newRow.Cell(21).Value = newAccountingElements[counter].VATSales;

                    newRow.Cell(22).Value = newAccountingElements[counter].SalesVAT;                    // 3000
                    newRow.Cell(23).Value = newAccountingElements[counter].SalesVATExempt;              // 3100

                    newRow.Cell(24).Value = newAccountingElements[counter].CostOfGoods;                 // 4005
                    newRow.Cell(25).Value = newAccountingElements[counter].CostForReselling;            // 4300
                    newRow.Cell(26).Value = newAccountingElements[counter].CostForSalary;               // 5000
                    newRow.Cell(27).Value = newAccountingElements[counter].CostForSalaryTax;            // 5400
                    newRow.Cell(28).Value = newAccountingElements[counter].CostForDepreciation;         // 6000
                    newRow.Cell(29).Value = newAccountingElements[counter].CostForShipping;             // 6100
                    newRow.Cell(30).Value = newAccountingElements[counter].CostForElectricity;          // 6340 
                    newRow.Cell(31).Value = newAccountingElements[counter].CostForToolsInventory;       // 6500
                    newRow.Cell(32).Value = newAccountingElements[counter].CostForMaintenance;          // 6695
                    newRow.Cell(33).Value = newAccountingElements[counter].CostForFacilities;           // 6800 

                    newRow.Cell(34).Value = newAccountingElements[counter].CostOfData;                  // 6810 
                    newRow.Cell(35).Value = newAccountingElements[counter].CostOfPhoneInternet;         // 6900
                    newRow.Cell(36).Value = newAccountingElements[counter].CostForTravelAndAllowance;   // 7140
                    newRow.Cell(37).Value = newAccountingElements[counter].CostOfAdvertising;           // 7330
                    newRow.Cell(38).Value = newAccountingElements[counter].CostOfOther;                 // 7700

                    newRow.Cell(39).Value = newAccountingElements[counter].FeesBank;                    // 7770
                    newRow.Cell(40).Value = newAccountingElements[counter].FeesPaypal;                  // 7780
                    newRow.Cell(41).Value = newAccountingElements[counter].FeesStripe;                  // 7785 

                    newRow.Cell(42).Value = newAccountingElements[counter].CostForEstablishment;        // 7790

                    newRow.Cell(43).Value = newAccountingElements[counter].IncomeFinance;               // 8099
                    newRow.Cell(44).Value = newAccountingElements[counter].CostOfFinance;               // 8199

                    newRow.Cell(45).Value = newAccountingElements[counter].Investments;                 // 1200
                    newRow.Cell(46).Value = newAccountingElements[counter].AccountsReceivable;          // 1500
                    newRow.Cell(47).Value = newAccountingElements[counter].PersonalWithdrawal;
                    newRow.Cell(48).Value = newAccountingElements[counter].PersonalDeposit;

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
            row.Cells("T", "U").Style.Fill.BackgroundColor = green;

            // set background color for investments, withdrawal and deposits
            var lightBlue = XLColor.FromArgb(0xC5D9F1);
            var lighterBlue = XLColor.FromArgb(0xEAF1FA);
            var blue = currentRow % 2 == 0 ? lightBlue : lighterBlue;
            row.Cells("AS", "AV").Style.Fill.BackgroundColor = blue;

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
            row.Cells("P", "AX").Style.NumberFormat.Format = "#,##0.00;[Red]-#,##0.00;";
            row.Cells("P", "AX").DataType = XLCellValues.Number;
        }

        static void SetExcelTableTotalsRowFunction(IXLTable table)
        {
            table.ShowTotalsRow = true;

            // set sum functions for each of the table columns 
            table.Field("Paypal").TotalsRowFunction = XLTotalsRowFunction.Sum;              // 1910
            table.Field("Stripe").TotalsRowFunction = XLTotalsRowFunction.Sum;              // 1915
            table.Field("Vipps").TotalsRowFunction = XLTotalsRowFunction.Sum;               // 1918
            table.Field("Bank").TotalsRowFunction = XLTotalsRowFunction.Sum;                // 1920

            table.Field("MVA Kjøp").TotalsRowFunction = XLTotalsRowFunction.Sum;
            table.Field("MVA Salg").TotalsRowFunction = XLTotalsRowFunction.Sum;

            table.Field("Salg mva-pliktig").TotalsRowFunction = XLTotalsRowFunction.Sum;                   // 3000
            table.Field("Salg avgiftsfritt").TotalsRowFunction = XLTotalsRowFunction.Sum;             // 3100

            table.Field("Varekostnad").TotalsRowFunction = XLTotalsRowFunction.Sum;                // 4005
            table.Field("Forbruk for videresalg").TotalsRowFunction = XLTotalsRowFunction.Sum;           // 4300
            table.Field("Lønn").TotalsRowFunction = XLTotalsRowFunction.Sum;              // 5000
            table.Field("Arb.giver avgift").TotalsRowFunction = XLTotalsRowFunction.Sum;           // 5400
            table.Field("Avskrivninger").TotalsRowFunction = XLTotalsRowFunction.Sum;        // 6000
            table.Field("Frakt").TotalsRowFunction = XLTotalsRowFunction.Sum;            // 6100
            table.Field("Strøm").TotalsRowFunction = XLTotalsRowFunction.Sum;         // 6340 
            table.Field("Verktøy inventar").TotalsRowFunction = XLTotalsRowFunction.Sum;      // 6500
            table.Field("Vedlikehold").TotalsRowFunction = XLTotalsRowFunction.Sum;         // 6695
            table.Field("Kontorkostnader").TotalsRowFunction = XLTotalsRowFunction.Sum;          // 6800 

            table.Field("Datakostnader").TotalsRowFunction = XLTotalsRowFunction.Sum;                 // 6810 
            table.Field("Telefon Internett").TotalsRowFunction = XLTotalsRowFunction.Sum;        // 6900
            table.Field("Reise og Diett").TotalsRowFunction = XLTotalsRowFunction.Sum;  // 7140
            table.Field("Reklamekostnader").TotalsRowFunction = XLTotalsRowFunction.Sum;          // 7330
            table.Field("Diverse annet").TotalsRowFunction = XLTotalsRowFunction.Sum;                // 7700

            table.Field("Gebyrer Bank").TotalsRowFunction = XLTotalsRowFunction.Sum;                   // 7770
            table.Field("Gebyrer Paypal").TotalsRowFunction = XLTotalsRowFunction.Sum;                 // 7780
            table.Field("Gebyrer Stripe").TotalsRowFunction = XLTotalsRowFunction.Sum;                 // 7785 

            table.Field("Etableringskostnader").TotalsRowFunction = XLTotalsRowFunction.Sum;       // 7790

            table.Field("Finansinntekter").TotalsRowFunction = XLTotalsRowFunction.Sum;              // 8099
            table.Field("Finanskostnader").TotalsRowFunction = XLTotalsRowFunction.Sum;              // 8199

            table.Field("Investeringer").TotalsRowFunction = XLTotalsRowFunction.Sum;                // 1200
            table.Field("Kundefordringer").TotalsRowFunction = XLTotalsRowFunction.Sum;         // 1500
            table.Field("Privat uttak").TotalsRowFunction = XLTotalsRowFunction.Sum;
            table.Field("Privat innskudd").TotalsRowFunction = XLTotalsRowFunction.Sum;

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
