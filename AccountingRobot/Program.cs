using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using FastExcel;
using CsvHelper;
using System.Globalization;
using CsvHelper.TypeConversion;
using CsvHelper.Configuration;
using System.Configuration;
using Newtonsoft.Json;
using System.Net;

namespace AccountingRobot
{
    class Program
    {
        static void Main(string[] args)
        {
            //ReadOberloOrders(@"C:\Users\pnerseth\Amazon Drive\Documents\Private\wazalo\regnskap\Oberlo Orders 2017-01-01-2017-12-04.xlsx");
            //ReadAliExpressOrders(@"C:\Users\pnerseth\Amazon Drive\Documents\Private\wazalo\regnskap\AliExpressOrders-2017-12-10_00-59.csv");

            string shopifyDomain = ConfigurationManager.AppSettings["ShopifyDomain"];
            string shopifyAPIKey = ConfigurationManager.AppSettings["ShopifyAPIKey"];
            string shopifyAPIPassword = ConfigurationManager.AppSettings["ShopifyAPIPassword"];
            int count = CountShopifyOrders(shopifyDomain, shopifyAPIKey, shopifyAPIPassword);
            Console.WriteLine("{0} shopify orders", count);
            //ReadShopifyOrders(shopifyDomain, shopifyAPIKey, shopifyAPIPassword);

            Console.ReadLine();
        }

        static int CountShopifyOrders(string shopifyDomain, string shopifyAPIKey, string shopifyAPIPassword)
        {
            // GET /admin/orders/count.json
            // Retrieve a count of all the orders

            string url = String.Format("https://{0}/admin/orders/count.json?financial_status=paid&status=any", shopifyDomain);

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

        static void ReadShopifyOrders(string shopifyDomain, string shopifyAPIKey, string shopifyAPIPassword)
        {
            // GET /admin/orders.json
            // Retrieve a list of Orders(OPEN Orders by default, use status=any for ALL orders)
            // GET /admin/orders/#{id}.json
            // Receive a single Order

            // financial_status=paid
            // status=any

            // By default that Orders API endpoint can give you a maximum of 50 orders. 
            // You can increase the limit to 250 orders by adding &limit=250 to the URL. 
            // If your query has more than 250 results then you can page through them 
            // by using the page URL parameter: https://help.shopify.com/api/reference/order
            // limit: Amount of results (default: 50)(maximum: 250)
            // page: Page to show, (default: 1)

            string url = String.Format("https://{0}/admin/orders.json?financial_status=paid&status=any", shopifyDomain);

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
                    long id = order.id;
                    string name = order.name;
                    string financialStatus = order.financial_status;
                    string gateway = order.gateway;
                    decimal totalPrice = order.total_price;
                    decimal totalTax = order.total_tax;
                    string customerName = order.customer.first_name;

                    Console.WriteLine("{0} {1} {2} {3}", id, name, totalPrice, customerName);
                }
            }

        }

        static void ReadShopifyOrderPage(string shopifyDomain, string shopifyAPIKey, string shopifyAPIPassword)
        {
            // https://ecommerce.shopify.com/c/shopify-apis-and-technology/t/paginate-api-results-113066
            // GET /admin/products.xml?limit=250&page=1
            // Use the /admin/products/count.json to get the count of all products. Then divide that number by 250 to get the total amount of pages.
            /*
            if count > 0
                page += count.divmod(250).first
                while page > 0
                    puts "Processing page #{page}"
                    products += ShopifyAPI::Product.all(:params => {:page => page, :limit => 250})
                    page -= 1
                end
            end
      */
        }

        static void ReadOberloOrders(string oberloOrdersFilesPath)
        {
            // Get the input file paths
            FileInfo inputFile = new FileInfo(oberloOrdersFilesPath);

            //Create a worksheet
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
                    var orderNumber = row.GetCellByColumnName("A").Value; // Order #
                    var createdDate = row.GetCellByColumnName("B").Value; // Created
                    var financialStatus = row.GetCellByColumnName("C").Value; // Financial status                    
                    var supplier = row.GetCellByColumnName("G").Value; // Supplier
                    var SKU = row.GetCellByColumnName("H").Value; // SKU
                    var productName = row.GetCellByColumnName("I").Value; // Product
                    var variant = row.GetCellByColumnName("J").Value; // Variant
                    var trackingNumber = row.GetCellByColumnName("L"); // Tracking number
                    var aliOrderNumber = row.GetCellByColumnName("M"); // Ali Order #
                    var customerName = row.GetCellByColumnName("N").Value; // Name
                    var customerAddress = row.GetCellByColumnName("O").Value; // Address
                    var customerAddress2 = row.GetCellByColumnName("P"); // Address2
                    var customerCity = row.GetCellByColumnName("Q").Value; // City
                    var customerZip = row.GetCellByColumnName("R").Value; // Zip
                    var orderNote = row.GetCellByColumnName("V"); // Order note
                    var orderState = row.GetCellByColumnName("W").Value; // Order state

                    var aliOrderNumberString = (aliOrderNumber != null ? aliOrderNumber.Value.ToString() : "");
                    var trackingNumberString = (trackingNumber != null ? trackingNumber.Value.ToString() : "");
                    var customerAddress2String = (customerAddress2 != null ? customerAddress2.Value.ToString() : "");
                    var orderNoteString = (orderNote != null ? orderNote.Value.ToString() : "");
                    var created = DateTime.ParseExact(createdDate.ToString(), "yyyy-MM-dd", CultureInfo.InvariantCulture);

                    Console.WriteLine("{0} {1:yyyy-MM-dd} {2} {3} {4}", orderNumber, created, aliOrderNumberString, SKU, customerName);
                }
            }
        }

        static void ReadAliExpressOrders(string aliExpressOrdersFilesPath)
        {
            using (TextReader fileReader = File.OpenText(aliExpressOrdersFilesPath))
            {
                fileReader.ReadLine(); // skip excel separator line

                using (var csvReader = new CsvReader(fileReader))
                {
                    csvReader.Configuration.Delimiter = ",";
                    csvReader.Configuration.HasHeaderRecord = true;
                    csvReader.Configuration.CultureInfo = CultureInfo.InvariantCulture;
                    csvReader.Configuration.RegisterClassMap<CsvMap>();

                    var aliExpressOrders = csvReader.GetRecords<AliExpressOrder>();

                    foreach (var aliExpressOrder in aliExpressOrders)
                    {
                        Console.WriteLine("{0} {1} {2} {3}", aliExpressOrder.OrderId, aliExpressOrder.OrderTime, aliExpressOrder.OrderAmount, aliExpressOrder.ContactName);
                    }
                }
            }

        }

        public class AliExpressOrder
        {
            public long OrderId { get; set; }
            public string OrderTime { get; set; }
            public string StoreName { get; set; }
            public string StoreUrl { get; set; }
            public decimal OrderAmount { get; set; }
            public string OrderLines { get; set; }
            public string ContactName { get; set; }
            public string ContactAddress { get; set; }
            public string ContactAddress2 { get; set; }
            public string ContactZipCode { get; set; }
        }

        public sealed class CsvMap : ClassMap<AliExpressOrder>
        {
            public CsvMap()
            {
                Map(m => m.OrderId);
                Map(m => m.OrderTime).TypeConverterOption.Format("HH:mm MMM. dd yyyy");
                Map(m => m.StoreName);
                Map(m => m.StoreUrl);
                Map(m => m.OrderAmount).TypeConverter<CustomCurrencyConverter>();
                Map(m => m.OrderLines);
                Map(m => m.ContactName);
                Map(m => m.ContactAddress);
                Map(m => m.ContactAddress2);
                Map(m => m.ContactZipCode);
            }
        }

        public class CustomCurrencyConverter : ITypeConverter
        {
            public object ConvertFromString(string text, IReaderRow row, MemberMapData memberMapData)
            {
                // convert string like "$ 19.80" to decimal         
                var numberFormat = new NumberFormatInfo();
                numberFormat.NegativeSign = "-";
                numberFormat.CurrencyDecimalSeparator = ".";
                numberFormat.CurrencyGroupSeparator = "";
                numberFormat.CurrencySymbol = "$ ";

                return decimal.Parse(text, NumberStyles.Currency, numberFormat);
            }

            public string ConvertToString(object value, IWriterRow row, MemberMapData memberMapData)
            {
                throw new NotImplementedException();
            }
        }
    }
}
