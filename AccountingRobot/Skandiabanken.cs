using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Threading;
using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace AccountingRobot
{
    public static class Skandiabanken
    {
        public static SkandiabankenBankStatement ReadBankStatement(string skandiabankenTransactionsFilePath)
        {
            var skandiabankenTransactions = new List<SkandiabankenTransaction>();

            var wb = new XLWorkbook(skandiabankenTransactionsFilePath);
            var ws = wb.Worksheet("Kontoutskrift");

            var startColumn = ws.Column(1);
            var firstCellFirstColumn = startColumn.FirstCellUsed();
            var lastCellFirstColumn = startColumn.LastCellUsed();
            var lastCellLastColumn = lastCellFirstColumn.WorksheetRow().AsRange().LastColumnUsed().LastCellUsed();

            // Get a range with the transaction data
            var transactionRange = ws.Range(firstCellFirstColumn, lastCellLastColumn).RangeUsed();

            // Treat the range as a table
            var transactionTable = transactionRange.AsTable();

            // Get the transactions
            foreach (var row in transactionTable.DataRange.Rows())
            {
                // BOKFØRINGSDATO	
                // RENTEDATO	
                // ARKIVREFERANSE	
                // TYPE	
                // TEKST	
                // UT FRA KONTO	
                // INN PÅ KONTO
                var skandiabankenTransaction = new SkandiabankenTransaction();
                skandiabankenTransaction.TransactionDate = row.Field(0).GetDateTime();
                skandiabankenTransaction.InterestDate = row.Field(1).GetDateTime();
                skandiabankenTransaction.ArchiveReference = row.Field(2).GetValue<long>();
                skandiabankenTransaction.Type = row.Field(3).GetString();
                skandiabankenTransaction.Text = row.Field(4).GetString();
                skandiabankenTransaction.OutAccount = row.Field(5).GetValue<decimal>();
                skandiabankenTransaction.InAccount = row.Field(6).GetValue<decimal>();

                // set account change
                decimal accountChange = skandiabankenTransaction.InAccount - skandiabankenTransaction.OutAccount; ;
                skandiabankenTransaction.AccountChange = accountChange;

                if (accountChange > 0)
                {
                    skandiabankenTransaction.AccountingType = SkandiabankenTransaction.AccountingTypeEnum.IncomeUnknown;
                }
                else
                {
                    skandiabankenTransaction.AccountingType = SkandiabankenTransaction.AccountingTypeEnum.CostUnknown;
                }

                skandiabankenTransactions.Add(skandiabankenTransaction);
            }

            // find the incoming and outgoing balance
            var incomingBalanceCell = ws.Cell(lastCellLastColumn.Address.RowNumber + 2, lastCellLastColumn.Address.ColumnNumber);
            var outgoingBalanceCell = ws.Cell(1, lastCellLastColumn.Address.ColumnNumber);
            decimal incomingBalance = incomingBalanceCell.GetValue<decimal>();
            decimal outgoingBalance = outgoingBalanceCell.GetValue<decimal>();
            var incomingBalanceLabelCell = ws.Cell(lastCellLastColumn.Address.RowNumber + 2, lastCellLastColumn.Address.ColumnNumber-2);
            var outgoingBalanceLabelCell = ws.Cell(1, lastCellLastColumn.Address.ColumnNumber-2);
            var incomingBalanceLabel = incomingBalanceLabelCell.GetString();
            var outgoingBalanceLabel = outgoingBalanceLabelCell.GetString();
            var incomingBalanceDate = ExcelUtils.GetDateFromBankStatementString(incomingBalanceLabel);
            var outgoingBalanceDate = ExcelUtils.GetDateFromBankStatementString(outgoingBalanceLabel);

            var bankStatment = new SkandiabankenBankStatement
            {
                Transactions = skandiabankenTransactions,
                IncomingBalanceDate = incomingBalanceDate,
                IncomingBalanceLabel = incomingBalanceLabel,
                IncomingBalance = incomingBalance,
                OutgoingBalanceDate = outgoingBalanceDate,
                OutgoingBalanceLabel = outgoingBalanceLabel,
                OutgoingBalance = outgoingBalance
            };

            return bankStatment;
        }

        public static bool DownloadBankStatement()
        {
            string cacheDir = ConfigurationManager.AppSettings["CacheDir"];
            string userDataDir = ConfigurationManager.AppSettings["UserDataDir"];

            // C:\Users\pnerseth\Downloads\97132735232_2017_01_01-2017_12_27.xlsx
            string downloadFolderPath = @"C:\Users\pnerseth\Downloads";

            string sbankenMobilePhone = "90156615";
            string sbankenBirthDate = "070374";

            var currentDate = DateTime.Now.Date;
            var firstDayOfTheYear = new DateTime(currentDate.Year, 1, 1);
            var yesterday = currentDate.AddDays(-1);

            //string sbankenCustomFromDate = string.Format("{0:dd.MM.yyyy}", firstDayOfTheYear);
            //string sbankenCustomToDate = string.Format("{0:dd.MM.yyyy}", yesterday);

            string userDataArgument = string.Format("user-data-dir={0}", userDataDir);

            // http://blog.hanxiaogang.com/2017-07-29-aliexpress/
            ChromeOptions options = new ChromeOptions();
            options.AddArguments(userDataArgument);
            options.AddArguments("--start-maximized");
            options.AddArgument("--log-level=3");
            IWebDriver driver = new ChromeDriver(options);

            driver.Navigate().GoToUrl("https://secure.sbanken.no/Authentication/BankIdMobile");

            var waitLoginPage = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            waitLoginPage.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));

            // login if login form is present
            if (SeleniumUtils.IsElementPresent(driver, By.Id("MobilePhone")))
            {
                // https://secure.sbanken.no/Authentication/BankIDMobile
                // input id MobilePhone - Mobilnummer (8 siffer)
                // input it BirthDate - Fødselsdato (ddmmåå)
                // submit value = Neste

                // login if login form is present
                if (SeleniumUtils.IsElementPresent(driver, By.XPath("//input[@id='MobilePhone']"))
                    && SeleniumUtils.IsElementPresent(driver, By.XPath("//input[@id='BirthDate']")))
                {
                    IWebElement mobilePhone = driver.FindElement(By.XPath("//input[@id='MobilePhone']"));
                    IWebElement birthDate = driver.FindElement(By.XPath("//input[@id='BirthDate']"));

                    mobilePhone.Clear();
                    mobilePhone.SendKeys(sbankenMobilePhone);

                    birthDate.Clear();
                    birthDate.SendKeys(sbankenBirthDate);

                    // use birth date field to submit form
                    birthDate.Submit();

                    var waitLoginIFrame = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                    waitLoginIFrame.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
                }
            }

            try
            {
                var waitMainPage = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                waitMainPage.Until(ExpectedConditions.UrlToBe("https://secure.sbanken.no/Home/Overview/Full"));
            }
            catch (WebDriverTimeoutException)
            {
                Console.WriteLine("Timeout - Logged in to Skandiabanken to late. Stopping.");
                return false;
            }

            // go to account statement
            driver.Navigate().GoToUrl("https://secure.sbanken.no/Home/AccountStatement?accountId=483c5027bafcf1d43e623a99d6a0e0e8");

            // //*[@id="expanded-filters-link"]
            // input SearchFilter_CustomFromDate (dd.mm.yyyyy)
            // input SearchFilter_CustomToDate (dd.mm.yyyyy)
            // submit value = Vis
            /*
            if (SeleniumUtils.IsElementPresent(driver, By.XPath("//input[@id='SearchFilter_CustomFromDate']"))
                && SeleniumUtils.IsElementPresent(driver, By.XPath("//input[@id='SearchFilter_CustomToDate']")))
            {
                IWebElement customFromDate = driver.FindElement(By.XPath("//input[@id='SearchFilter_CustomFromDate']"));
                IWebElement customToDate = driver.FindElement(By.XPath("//input[@id='SearchFilter_CustomToDate']"));

                customFromDate.Clear();
                customFromDate.SendKeys(sbankenCustomFromDate);

                customToDate.Clear();
                customToDate.SendKeys(sbankenCustomToDate);

                // use birth date field to submit form
                customToDate.Submit();

                var waitLoginIFrame = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                waitLoginIFrame.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
            }
            */

            // download account statement
            string accountStatementDownload = string.Format("https://secure.sbanken.no/Home/AccountStatement/ViewExcel?AccountId=483c5027bafcf1d43e623a99d6a0e0e8&CustomFromDate={0:dd.MM.yyyy}&CustomToDate={1:dd.MM.yyyy}&FromDate=CustomPeriod&Incoming=", firstDayOfTheYear, yesterday);
            driver.Navigate().GoToUrl(accountStatementDownload);

            var waitExcel = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            waitExcel.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));

            string accountStatementFileName = string.Format("97132735232_{0:yyyy_MM_dd}-{1:yyyy_MM_dd}.xlsx", firstDayOfTheYear, yesterday);
            string accountStatementFilePath = Path.Combine(downloadFolderPath, accountStatementFileName);

            // wait until file has downloaded
            for (var i = 0; i < 30; i++)
            {
                if (File.Exists(accountStatementFilePath)) { break; }
                Thread.Sleep(1000);
            }
            var length = new FileInfo(accountStatementFilePath).Length;
            for (var i = 0; i < 30; i++)
            {
                Thread.Sleep(1000);
                var newLength = new FileInfo(accountStatementFilePath).Length;
                if (newLength == length && length != 0) { break; }
                length = newLength;
            }
            driver.Close();

            // determine path
            Console.Out.WriteLine("Successfully downloaded skandiabanken account statement excel file {0}", accountStatementFilePath);

            // moving file to right place
            string destFilePath = Path.Combine(cacheDir, accountStatementFileName);

            // To copy a folder's contents to a new location:
            // Create a new target folder, if necessary.
            if (!Directory.Exists(cacheDir))
            {
                Directory.CreateDirectory(cacheDir);
            }

            // Move file to another location
            File.Move(accountStatementFilePath, destFilePath);

            return true;
        }
    }

    public class SkandiabankenBankStatement
    {
        public List<SkandiabankenTransaction> Transactions { get; set; }
        public DateTime IncomingBalanceDate { get; set; }
        public string IncomingBalanceLabel { get; set; }
        public decimal IncomingBalance { get; set; }
        public DateTime OutgoingBalanceDate { get; set; }
        public string OutgoingBalanceLabel { get; set; }
        public decimal OutgoingBalance { get; set; }
    }
}
