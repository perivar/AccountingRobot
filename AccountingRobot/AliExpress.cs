﻿using System;
using System.Text;
using System.Text.RegularExpressions;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.IO;
using CsvHelper;
using System.Configuration;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using AccountingRobot;

namespace AliOrderScraper
{
    public static class AliExpress
    {
        public static List<AliExpressOrder> GetLatestAliExpressOrders(bool forceUpdate = false)
        {
            string userDataDir = ConfigurationManager.AppSettings["UserDataDir"];
            string cacheDir = ConfigurationManager.AppSettings["CacheDir"];
            string cacheFileNamePrefix = "AliExpress Orders";

            var lastCacheFile = Utils.FindLastCacheFile(cacheDir, cacheFileNamePrefix);

            var currentDate = DateTime.Now.Date;
            var firstDayOfTheYear = new DateTime(currentDate.Year, 1, 1);
            var lastDayOfTheYear = new DateTime(currentDate.Year, 12, 31);

            // check if we have a cache file
            DateTime from = default(DateTime);
            DateTime to = default(DateTime);

            // if the cache file object has values
            if (!lastCacheFile.Equals(default(KeyValuePair<DateTime, string>)))
            {
                from = lastCacheFile.Key.Date;
                to = currentDate;

                // if the from date is today, then we already have an updated file so use cache
                if (from.Equals(to))
                {
                    // use latest cache file (or force an update)
                    return GetAliExpressOrders(lastCacheFile.Value, userDataDir, from, forceUpdate);
                }
                else if (from != firstDayOfTheYear)
                {
                    // we have to combine two files:
                    // the original cache file and the new transactions file
                    Console.Out.WriteLine("Finding AliExpress Orders from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
                    var newAliExpressOrders = ScrapeAliExpressOrders(userDataDir, from);
                    var originalAliExpressOrders = Utils.ReadCacheFile<AliExpressOrder>(lastCacheFile.Value);

                    // copy all the original AliExpress orders into a new file, except entries that are 
                    // from the from date or newer
                    var updatedAliExpressOrders = originalAliExpressOrders.Where(p => p.OrderTime < from).ToList();

                    // and add the new orders to beginning of list
                    updatedAliExpressOrders.InsertRange(0, newAliExpressOrders);

                    // and store to new file
                    string newCacheFilePath = Path.Combine(cacheDir, string.Format("{0}-{1:yyyy-MM-dd}-{2:yyyy-MM-dd}.csv", cacheFileNamePrefix, firstDayOfTheYear, to));
                    using (var sw = new StreamWriter(newCacheFilePath))
                    {
                        var csvWriter = new CsvWriter(sw);
                        csvWriter.Configuration.Delimiter = ",";
                        csvWriter.Configuration.HasHeaderRecord = true;
                        csvWriter.Configuration.CultureInfo = CultureInfo.InvariantCulture;

                        csvWriter.WriteRecords(updatedAliExpressOrders);
                    }

                    Console.Out.WriteLine("Successfully wrote file to {0}", newCacheFilePath);
                    return updatedAliExpressOrders;
                }
            }
            else
            {
                // find all from beginning of year until now
                from = firstDayOfTheYear;
                to = currentDate;
            }

            // get updated transactions (or from cache file if update is forced)
            string cacheFilePath = Path.Combine(cacheDir, string.Format("{0}-{1:yyyy-MM-dd}-{2:yyyy-MM-dd}.csv", cacheFileNamePrefix, from, to));
            return GetAliExpressOrders(cacheFilePath, userDataDir, from);
        }

        public static List<AliExpressOrder> GetAliExpressOrders(string cacheFilePath, string userDataDir, DateTime from, bool forceUpdate = false)
        {
            var cachedAliExpressOrders = Utils.ReadCacheFile<AliExpressOrder>(cacheFilePath, forceUpdate);
            if (cachedAliExpressOrders != null && cachedAliExpressOrders.Count() > 0)
            {
                Console.Out.WriteLine("Using cache file {0}.", cacheFilePath);
                return cachedAliExpressOrders;
            }
            else
            {
                var AliExpressOrders = ScrapeAliExpressOrders(userDataDir, from);

                using (var sw = new StreamWriter(cacheFilePath))
                {
                    var csvWriter = new CsvWriter(sw);
                    csvWriter.Configuration.Delimiter = ",";
                    csvWriter.Configuration.HasHeaderRecord = true;
                    csvWriter.Configuration.CultureInfo = CultureInfo.InvariantCulture;

                    csvWriter.WriteRecords(AliExpressOrders);
                }

                Console.Out.WriteLine("Successfully wrote file to {0}", cacheFilePath);
                return AliExpressOrders;
            }
        }

        static List<AliExpressOrder> ScrapeAliExpressOrders(string userDataDir, DateTime from)
        {
            var aliExpressOrders = new List<AliExpressOrder>();

            string userDataArgument = string.Format("user-data-dir={0}", userDataDir);

            // http://blog.hanxiaogang.com/2017-07-29-aliexpress/
            ChromeOptions options = new ChromeOptions();
            options.AddArguments(userDataArgument);
            options.AddArguments("--start-maximized");
            options.AddArgument("--log-level=3");
            //options.AddArguments("--ignore-certificate-errors");
            //options.AddArguments("--ignore-ssl-errors");
            IWebDriver driver = new ChromeDriver(options);
            driver.Navigate().GoToUrl("https://login.aliexpress.com");

            try
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                wait.Until(ExpectedConditions.UrlToBe("https://www.aliexpress.com/"));
            }
            catch (WebDriverTimeoutException)
            {
                Console.WriteLine("Timeout - Logged in to AliExpress to late. Stopping.");
                return aliExpressOrders;
            }

            // go to order list
            driver.Navigate().GoToUrl("https://trade.aliexpress.com/orderList.htm");

            // identify how many pages on order page (1/20)
            var tuple = GetAliExpressOrderPageNumber(driver);
            int curPage = tuple.Item1;
            int numPages = tuple.Item2;
            Console.WriteLine("Found {0} Pages", numPages);

            // scrape one and one page
            for (int i = 1; i <= numPages; i++)
            {
                // if this method returns false, it means we have reached the from date
                if (!ScrapeAliExpressOrderPage(aliExpressOrders, driver, i, from))
                {
                    break;
                }
            }

            return aliExpressOrders;
        }

        static bool ScrapeAliExpressOrderPage(List<AliExpressOrder> aliExpressOrders, IWebDriver driver, int curPage, DateTime from)
        {
            // change page
            driver.FindElement(By.CssSelector("input[id$='gotoPageNum']")).SendKeys(curPage.ToString());
            driver.FindElement(By.XPath("//*[@id='btnGotoPageNum']")).Click();

            // check that we in fact got to the right page
            var tuple = GetAliExpressOrderPageNumber(driver);
            int newPage = tuple.Item1;
            int numPages = tuple.Item2;

            if (curPage != newPage) return false;

            // scrape until we get a false back
            Console.WriteLine("Reading Page {0} of {1} Pages", curPage, numPages);
            return ScrapeAliExpressOrderPageEntry(aliExpressOrders, driver, curPage, from);
        }

        static bool ScrapeAliExpressOrderPageEntry(List<AliExpressOrder> aliExpressOrders, IWebDriver driver, int curPage, DateTime from)
        {
            var orderEntries = driver.FindElements(By.XPath("//tbody[contains(@class, 'order-item-wraper ')]"));
            Console.WriteLine("Found {0} orders on page {1}", orderEntries.Count, curPage);

            int pageCount = 1;
            foreach (var orderEntry in orderEntries)
            {
                Console.WriteLine("Reading order number {0}", pageCount++);

                var orderId = orderEntry.FindElement(By.XPath("tr[@class='order-head']/td[@class='order-info']/p[@class='first-row']/span[@class='info-body']")).Text;
                var orderTime = orderEntry.FindElement(By.XPath("tr[@class='order-head']/td[@class='order-info']/p[@class='second-row']/span[@class='info-body']")).Text;

                var storeName = orderEntry.FindElement(By.XPath("tr[@class='order-head']/td[@class='store-info']/p[@class='first-row']/span[@class='info-body']")).Text;
                var storeUrl = orderEntry.FindElement(By.XPath("tr[@class='order-head']/td[@class='store-info']/p[@class='second-row']/a")).GetAttribute("href");

                var orderAmount = orderEntry.FindElement(By.XPath("tr[@class='order-head']/td[@class='order-amount']/div[@class='amount-body']/p[@class='amount-num']")).Text;

                Console.WriteLine("Order no {0} was ordered on the {1}\n{2}, {3}, {4}", orderId, orderTime, storeName, storeUrl, orderAmount);

                var aliExpressOrder = new AliExpressOrder();
                aliExpressOrder.OrderId = long.Parse(orderId);
                aliExpressOrder.OrderTime = DateTime.ParseExact(orderTime, "HH:mm MMM. dd yyyy", CultureInfo.InvariantCulture);

                // if we have reached a day before the from date, stop
                if (aliExpressOrder.OrderTime.Date <= from.AddDays(-1).Date)
                {
                    Console.WriteLine("Reached the from date. (Found order from {0:dd.MM.yyyy} <= {1:dd.MM.yyyy})", aliExpressOrder.OrderTime.Date, from.AddDays(-1).Date);
                    return false;
                }

                aliExpressOrder.StoreName = storeName;
                aliExpressOrder.StoreUrl = storeUrl;
                aliExpressOrder.OrderAmount = ConvertDecimalFromString(orderAmount);

                // for each order line
                var orderLines = orderEntry.FindElements(By.XPath("tr[@class='order-body']"));

                // add all orderlines to a string
                int orderLineCount = 1;
                bool first = true;
                var builder = new StringBuilder(); // initially empty
                foreach (var orderLine in orderLines)
                {
                    // append newline after each line
                    if (first)
                    {
                        first = false;
                    }
                    else
                    {
                        builder.Append("\n");
                    }

                    var productTitleElement = orderLine.FindElement(By.XPath("td[@class='product-sets']/div[@class='product-right']/p[@class='product-title']/a")); ;
                    var productId = productTitleElement.GetAttribute("productId");
                    var productTitle = productTitleElement.Text;

                    var productAmount = orderLine.FindElement(By.XPath("td[@class='product-sets']/div[@class='product-right']/p[@class='product-amount']")).Text;
                    var productProperty = orderLine.FindElement(By.XPath("td[@class='product-sets']/div[@class='product-right']/p[@class='product-property']")).Text;

                    Console.WriteLine("{0}. [{1}] {2}\n{3} {4}", orderLineCount, productId, productTitle, productProperty, productAmount);
                    builder.AppendFormat("{0}. [{1}] {2}\n{3} {4}", orderLineCount, productId, productTitle, productProperty, productAmount);
                    orderLineCount++;
                }
                aliExpressOrder.OrderLines = builder.ToString();

                // read order contact information (buyer)
                GetAliExpressContactFromOrder(aliExpressOrder, driver, orderId);

                // add to list
                aliExpressOrders.Add(aliExpressOrder);

                // new line
                Console.WriteLine();
            }

            return true;
        }

        static void GetAliExpressContactFromOrder(AliExpressOrder aliExpressOrder, IWebDriver driver, string orderId)
        {
            // https://trade.aliexpress.com/order_detail.htm?orderId=81495464493633

            // open a new tab and set the context
            var chromeDriver = (ChromeDriver)driver;

            // save a reference to our original tab's window handle
            var originalTabInstance = chromeDriver.CurrentWindowHandle;

            // execute some JavaScript to open a new window
            chromeDriver.ExecuteScript("window.open();");

            // save a reference to our new tab's window handle, this would be the last entry in the WindowHandles collection
            var newTabInstance = chromeDriver.WindowHandles[driver.WindowHandles.Count - 1];

            // switch our WebDriver to the new tab's window handle
            chromeDriver.SwitchTo().Window(newTabInstance);

            // lets navigate to a web site in our new tab
            string url = String.Format("https://trade.aliexpress.com/order_detail.htm?orderId={0}", orderId);
            driver.Navigate().GoToUrl(url);

            // find contact information
            // example:
            // <li><label> Contact Name :</label><span i18entitle = 'Contact Name' class="i18ncopy">Reidar Krogsaeter</span>
            // find the span element with correct i18entitle contained within the li that contains a label whose string value contains the substring Contact Name
            string contactName = driver.FindElement(By.XPath("//li[label[contains(., 'Contact Name')]]/span[@i18entitle='Contact Name']")).Text;
            string contactAddress = driver.FindElement(By.XPath("//li[label[contains(., 'Address')]]/span[@i18entitle='Address']")).Text;
            string contactZipCode = driver.FindElement(By.XPath("//li[label[contains(., 'Zip Code')]]/span[@i18entitle='Zip Code']")).Text;
            // example 2:
            // find the first following li element after the address and extract the span that has class i18ncopy
            string contactAddress2 = driver.FindElement(By.XPath("//li[label[contains(., 'Address')]]/following-sibling::li[1]/span[@class='i18ncopy']")).Text;

            Console.WriteLine("Contact {0} {1} {2} {3}", contactName, contactAddress, contactAddress2, contactZipCode);

            aliExpressOrder.ContactName = contactName;
            aliExpressOrder.ContactAddress = contactAddress;
            aliExpressOrder.ContactAddress2 = contactAddress2;
            aliExpressOrder.ContactZipCode = contactZipCode;

            // now lets close our new tab
            chromeDriver.ExecuteScript("window.close();");

            // and switch our WebDriver back to the original tab's window handle
            chromeDriver.SwitchTo().Window(originalTabInstance);

            // and have our WebDriver focus on the main document in the page to send commands to 
            chromeDriver.SwitchTo().DefaultContent();
        }

        static Tuple<int, int> GetAliExpressOrderPageNumber(IWebDriver driver)
        {
            int curPage = 0;
            int numPages = 0;
            var pageLabel = driver.FindElement(By.CssSelector("label[class$='ui-label']")).Text;
            Regex regexObj = new Regex(@"(\d+)/(\d+)", RegexOptions.IgnoreCase);
            Match matchResults = regexObj.Match(pageLabel);
            if (matchResults.Success)
            {
                curPage = int.Parse(matchResults.Groups[1].Value);
                numPages = int.Parse(matchResults.Groups[2].Value);
                return new Tuple<int, int>(curPage, numPages);
            }

            return null;
        }

        static decimal ConvertDecimalFromString(string text)
        {
            // convert string like "$ 19.80" to decimal         
            var numberFormat = new NumberFormatInfo();
            numberFormat.NegativeSign = "-";
            numberFormat.CurrencyDecimalSeparator = ".";
            numberFormat.CurrencyGroupSeparator = "";
            numberFormat.CurrencySymbol = "$ ";

            return decimal.Parse(text, NumberStyles.Currency, numberFormat);
        }

        public static List<AliExpressOrderGroup> CombineOrders(List<AliExpressOrder> aliExpressOrders)
        {
            var query = (from o in aliExpressOrders
                         group o by new { o.OrderTime, o.ContactName }
             into grp
                         select new AliExpressOrderGroup()
                         {
                             OrderTime = grp.Key.OrderTime,
                             ContactName = grp.Key.ContactName,
                             OrderAmount = grp.Sum(o => o.OrderAmount),
                             Children = grp.ToList()
                         }).ToList();

            return query;
        }
    }

    public class AliExpressOrderGroup
    {
        public DateTime OrderTime { get; set; }
        public decimal OrderAmount { get; set; }
        public string ContactName { get; set; }
        public List<AliExpressOrder> Children { get; set; }

        public override string ToString()
        {
            return string.Format("{0} {1:dd.MM} {2} {3}", string.Join(", ", Array.ConvertAll(Children.ToArray(), i => i.OrderId)), OrderTime, OrderAmount.ToString("C", new CultureInfo("en-US")), ContactName);
        }
    }
}
