using System;
using System.Collections.Generic;
using System.IO;
using FastExcel;
using CsvHelper;
using System.Globalization;
using System.Configuration;
using Newtonsoft.Json;
using System.Net;
using System.Linq;

namespace AccountingRobot
{
    partial class Program
    {
        static void Main(string[] args)
        {
            var oberloOrders = ReadOberloOrders(@"C:\Users\pnerseth\Amazon Drive\Documents\Private\wazalo\regnskap\Oberlo Orders 2017-01-01-2017-12-04.xlsx");
            foreach (var oberloOrder in oberloOrders)
            {
                Console.WriteLine("{0}", oberloOrder);
            }

            var skandiabankenTransactions = ReadSkandiabankenTransactions(@"C:\Users\pnerseth\Amazon Drive\Documents\Private\wazalo\regnskap\97132735232_2017_01_01-2017_12_10.xlsx");
            foreach (var skandiabankenTransaction in skandiabankenTransactions)
            {
                Console.WriteLine("{0}", skandiabankenTransaction);
            }

            var aliExpressOrders = ReadAliExpressOrders(@"C:\Users\pnerseth\Amazon Drive\Documents\Private\wazalo\regnskap\AliExpressOrders-2017-12-10_00-59.csv");
            foreach (var aliExpressOrder in aliExpressOrders)
            {
                Console.WriteLine("{0}", aliExpressOrder);
            }

            string shopifyDomain = ConfigurationManager.AppSettings["ShopifyDomain"];
            string shopifyAPIKey = ConfigurationManager.AppSettings["ShopifyAPIKey"];
            string shopifyAPIPassword = ConfigurationManager.AppSettings["ShopifyAPIPassword"];
            //string querySuffix = "financial_status=paid&status=any";
            string querySuffix = "status=any";
            int totalShopifyOrders = CountShopifyOrders(shopifyDomain, shopifyAPIKey, shopifyAPIPassword, querySuffix);
            Console.WriteLine("{0} shopify orders", totalShopifyOrders);

            var shopifyOrders = ReadShopifyOrders(shopifyDomain, shopifyAPIKey, shopifyAPIPassword, totalShopifyOrders, querySuffix);
            foreach (var order in shopifyOrders)
            {
                Console.WriteLine(order);
            }

            Console.ReadLine();
        }

        static List<SkandiabankenTransaction> ReadSkandiabankenTransactions(string skandiabankenTransactionsFilePath)
        {
            var skandiabankenTransactions = new List<SkandiabankenTransaction>();

            // Get the input file paths
            FileInfo inputFile = new FileInfo(skandiabankenTransactionsFilePath);

            // Create a worksheet
            Worksheet worksheet = null;

            // Create an instance of Fast Excel
            using (var fastExcel = new FastExcel.FastExcel(inputFile, true))
            {
                // Read the rows using worksheet name
                string worksheetName = "Kontoutskrift";
                worksheet = fastExcel.Read(worksheetName);

                Console.WriteLine("Reading worksheet {0} ...", worksheetName);

                // skip the three first rows since they only contain incoming balance and headers
                foreach (var row in worksheet.Rows.Skip(3))
                {
                    // read value rows
                    // BOKFØRINGSDATO	
                    // RENTEDATO	
                    // ARKIVREFERANSE	
                    // TYPE	
                    // TEKST	
                    // UT FRA KONTO	
                    // INN PÅ KONTO
                    var tmpValue = row.GetCellByColumnName("A").Value;

                    // if the first column (BOKFØRINGSDATO) field is empty we have reached the end 
                    // or not yet the start: The start is dealt with with the worksheet.Rows.Skip(3) command 
                    if (tmpValue.Equals(""))
                    {
                        break;
                    }
                    var transactionDateString = tmpValue.ToString();
                    var interestDateString = row.GetCellByColumnName("B").Value.ToString();
                    var archiveReferenceString = row.GetCellByColumnName("C").Value.ToString();
                    var type = row.GetCellByColumnName("D").Value.ToString();
                    var text = row.GetCellByColumnName("E").Value.ToString();
                    var outAccountString = row.GetCellByColumnName("F").Value.ToString();
                    var inAccountString = row.GetCellByColumnName("G").Value.ToString();

                    // convert to correct types
                    var transactionDate = GetDateFromExcelDateInt(transactionDateString);
                    var interestDate = GetDateFromExcelDateInt(interestDateString);
                    var archiveReference = long.Parse(archiveReferenceString);
                    decimal outAccount = GetDecimalFromExcelCurrencyString(outAccountString);
                    decimal inAccount = GetDecimalFromExcelCurrencyString(inAccountString);

                    // set account change
                    decimal accountChange = inAccount - outAccount;

                    var skandiabankenTransaction = new SkandiabankenTransaction();
                    skandiabankenTransaction.TransactionDate = transactionDate;
                    skandiabankenTransaction.InterestDate = interestDate;
                    skandiabankenTransaction.ArchiveReference = archiveReference;
                    skandiabankenTransaction.Type = type;
                    skandiabankenTransaction.Text = text;
                    skandiabankenTransaction.OutAccount = outAccount;
                    skandiabankenTransaction.InAccount = inAccount;
                    skandiabankenTransaction.AccountChange = accountChange;

                    skandiabankenTransactions.Add(skandiabankenTransaction);
                }
            }

            return skandiabankenTransactions;
        }

        /// <summary>
        /// Convert from excel date int string to actual date
        /// E.g. 39938 gets converted to 05/05/2009
        /// </summary>
        /// <param name="dateIntString">int string like 39938</param>
        /// <returns>datetime object (like 05/05/2009)</returns>
        static DateTime GetDateFromExcelDateInt(string dateIntString)
        {
            try
            {
                double d = double.Parse(dateIntString);
                DateTime conv = DateTime.FromOADate(d);
                return conv;
            }
            catch (Exception)
            {
                return DateTime.MinValue;
            }
        }

        /// <summary>
        /// Convert from excel decimal string to a decimal
        /// </summary>
        /// <param name="currencyString">currency string like 133.3</param>
        /// <returns>decimal like 133.3</returns>
        static decimal GetDecimalFromExcelCurrencyString(string currencyString)
        {
            //return Convert.ToDecimal(currencyString, CultureInfo.GetCultureInfo("no"));
            return Convert.ToDecimal(currencyString, CultureInfo.InvariantCulture);
        }

        static int CountShopifyOrders(string shopifyDomain, string shopifyAPIKey, string shopifyAPIPassword, string querySuffix)
        {
            // GET /admin/orders/count.json
            // Retrieve a count of all the orders

            string url = String.Format("https://{0}/admin/orders/count.json?{1}", shopifyDomain, querySuffix);

            using (var client = new WebClient())
            {
                // make sure we read in utf8
                client.Encoding = System.Text.Encoding.UTF8;

                // have to use the header field since normal GET url doesn't work, i.e.
                // string url = String.Format("https://{0}:{1}@{2}/admin/orders.json", shopifyAPIKey, shopifyAPIPassword, shopifyDomain);
                // https://stackoverflow.com/questions/28177871/shopify-and-private-applications
                client.Headers.Add("X-Shopify-Access-Token", shopifyAPIPassword);
                string json = client.DownloadString(url);

                // parse json
                dynamic jsonDe = JsonConvert.DeserializeObject(json);

                return jsonDe.count;
            }
        }

        static void ReadShopifyOrdersPage(List<ShopifyOrder> shopifyOrders, string shopifyDomain, string shopifyAPIKey, string shopifyAPIPassword, int limit, int page, string querySuffix)
        {
            // GET /admin/orders.json?limit=250&page=1
            // Retrieve a list of Orders(OPEN Orders by default, use status=any for ALL orders)

            // GET /admin/orders/#{id}.json
            // Receive a single Order

            // parameters:
            // financial_status=paid
            // status=any

            // By default that Orders API endpoint can give you a maximum of 50 orders. 
            // You can increase the limit to 250 orders by adding &limit=250 to the URL. 
            // If your query has more than 250 results then you can page through them 
            // by using the page URL parameter: https://help.shopify.com/api/reference/order
            // limit: Amount of results (default: 50)(maximum: 250)
            // page: Page to show, (default: 1)

            string url = String.Format("https://{0}/admin/orders.json?limit={1}&page={2}&{3}", shopifyDomain, limit, page, querySuffix);

            using (var client = new WebClient())
            {
                // make sure we read in utf8
                client.Encoding = System.Text.Encoding.UTF8;

                // have to use the header field since normal GET url doesn't work, i.e.
                // string url = String.Format("https://{0}:{1}@{2}/admin/orders.json", shopifyAPIKey, shopifyAPIPassword, shopifyDomain);
                // https://stackoverflow.com/questions/28177871/shopify-and-private-applications
                client.Headers.Add("X-Shopify-Access-Token", shopifyAPIPassword);
                string json = client.DownloadString(url);

                // parse json
                dynamic jsonDe = JsonConvert.DeserializeObject(json);

                foreach (var order in jsonDe.orders)
                {
                    var shopifyOrder = new ShopifyOrder();
                    shopifyOrder.Id = order.id;
                    shopifyOrder.Name = order.name;
                    shopifyOrder.FinancialStatus = order.financial_status;
                    string fulfillmentStatusTmp = order.fulfillment_status;
                    fulfillmentStatusTmp = (fulfillmentStatusTmp == null ? "unfulfilled" : fulfillmentStatusTmp);
                    shopifyOrder.FulfillmentStatus = fulfillmentStatusTmp;
                    shopifyOrder.Gateway = order.gateway;
                    shopifyOrder.TotalPrice = order.total_price;
                    shopifyOrder.TotalTax = order.total_tax;
                    shopifyOrder.CustomerName = string.Format("{0} {1}", order.customer.first_name, order.customer.last_name);

                    shopifyOrders.Add(shopifyOrder);
                }
            }
        }

        static List<ShopifyOrder> ReadShopifyOrders(string shopifyDomain, string shopifyAPIKey, string shopifyAPIPassword, int totalShopifyOrders, string querySuffix)
        {
            // https://ecommerce.shopify.com/c/shopify-apis-and-technology/t/paginate-api-results-113066
            // Use the /admin/products/count.json to get the count of all products. 
            // Then divide that number by 250 to get the total amount of pages.

            // the web api requires a pagination to read in all orders
            // max orders per page is 250

            var shopifyOrders = new List<ShopifyOrder>();

            int limit = 250;
            if (totalShopifyOrders > 0)
            {
                int numPages = (int)Math.Ceiling((double)totalShopifyOrders / limit);
                for (int i = 1; i <= numPages; i++)
                {
                    ReadShopifyOrdersPage(shopifyOrders, shopifyDomain, shopifyAPIKey, shopifyAPIPassword, limit, i, querySuffix);
                }
            }

            return shopifyOrders;
        }

        static List<OberloOrder> ReadOberloOrders(string oberloOrdersFilePath)
        {
            var oberloOrders = new List<OberloOrder>();

            // Get the input file paths
            FileInfo inputFile = new FileInfo(oberloOrdersFilePath);

            // Create a worksheet
            Worksheet worksheet = null;

            // Create an instance of Fast Excel
            using (var fastExcel = new FastExcel.FastExcel(inputFile, true))
            {
                // Read the rows using worksheet name
                string worksheetName = "Orders";
                worksheet = fastExcel.Read(worksheetName, 1);

                Console.WriteLine("Reading worksheet {0} ...", worksheetName);

                bool first = true;
                foreach (var row in worksheet.Rows)
                {
                    // skip header row
                    if (first)
                    {
                        first = false;
                        continue;
                    }

                    // read value rows
                    var orderNumber = row.GetCellByColumnName("A").Value.ToString(); // Order #
                    var createdDate = row.GetCellByColumnName("B").Value.ToString(); // Created
                    var financialStatus = row.GetCellByColumnName("C").Value.ToString(); // Financial status                    
                    var fulfillmentStatus = row.GetCellByColumnName("D").Value.ToString(); // Fulfillment status                    
                    var supplier = row.GetCellByColumnName("G").Value.ToString(); // Supplier
                    var SKU = row.GetCellByColumnName("H").Value.ToString(); // SKU
                    var productName = row.GetCellByColumnName("I").Value.ToString(); // Product
                    var variant = row.GetCellByColumnName("J").Value.ToString(); // Variant
                    var quantity = row.GetCellByColumnName("K").Value.ToString(); // Quantity
                    var trackingNumber = row.GetCellByColumnName("L"); // Tracking number
                    var aliOrderNumber = row.GetCellByColumnName("M"); // Ali Order #
                    var customerName = row.GetCellByColumnName("N").Value.ToString(); // Name
                    var customerAddress = row.GetCellByColumnName("O").Value.ToString(); // Address
                    var customerAddress2 = row.GetCellByColumnName("P"); // Address2
                    var customerCity = row.GetCellByColumnName("Q").Value.ToString(); // City
                    var customerZip = row.GetCellByColumnName("R").Value.ToString(); // Zip
                    var orderNote = row.GetCellByColumnName("V"); // Order note
                    var orderState = row.GetCellByColumnName("W").Value.ToString(); // Order state

                    var created = DateTime.ParseExact(createdDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                    var trackingNumberString = (trackingNumber != null ? trackingNumber.Value.ToString() : "");
                    var aliOrderNumberString = (aliOrderNumber != null ? aliOrderNumber.Value.ToString() : "");
                    var customerAddress2String = (customerAddress2 != null ? customerAddress2.Value.ToString() : "");
                    var orderNoteString = (orderNote != null ? orderNote.Value.ToString() : "");


                    var oberloOrder = new OberloOrder();
                    oberloOrder.OrderNumber = orderNumber;
                    oberloOrder.CreatedDate = created;
                    oberloOrder.FinancialStatus = financialStatus;
                    oberloOrder.FulfillmentStatus = fulfillmentStatus;
                    oberloOrder.Supplier = supplier;
                    oberloOrder.SKU = SKU;
                    oberloOrder.ProductName = productName;
                    oberloOrder.Variant = variant;
                    oberloOrder.Quantity = int.Parse(quantity);
                    oberloOrder.TrackingNumber = trackingNumberString;
                    oberloOrder.AliOrderNumber = aliOrderNumberString;
                    oberloOrder.CustomerName = customerName;
                    oberloOrder.CustomerAddress = customerAddress;
                    oberloOrder.CustomerAddress2 = customerAddress2String;
                    oberloOrder.CustomerCity = customerCity;
                    oberloOrder.CustomerZip = customerZip;
                    oberloOrder.OrderNote = orderNoteString;
                    oberloOrder.OrderState = orderState;

                    oberloOrders.Add(oberloOrder);
                }
            }

            return oberloOrders;
        }

        static List<AliExpressOrder> ReadAliExpressOrders(string aliExpressOrdersFilePath)
        {
            using (TextReader fileReader = File.OpenText(aliExpressOrdersFilePath))
            {
                fileReader.ReadLine(); // skip excel separator line

                using (var csvReader = new CsvReader(fileReader))
                {
                    csvReader.Configuration.Delimiter = ",";
                    csvReader.Configuration.HasHeaderRecord = true;
                    csvReader.Configuration.CultureInfo = CultureInfo.InvariantCulture;
                    csvReader.Configuration.RegisterClassMap<CsvMap>();

                    return csvReader.GetRecords<AliExpressOrder>().ToList<AliExpressOrder>();

                    /*
                    var aliExpressOrders = csvReader.GetRecords<AliExpressOrder>();

                    foreach (var aliExpressOrder in aliExpressOrders)
                    {
                        Console.WriteLine("{0}", aliExpressOrder);
                    }
                    */
                }
            }
        }
    }
}
