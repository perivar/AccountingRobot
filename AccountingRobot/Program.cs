using System;
using System.Configuration;

namespace AccountingRobot
{
    partial class Program
    {
        static void Main(string[] args)
        {
            var oberloOrders = Oberlo.ReadOrders(@"C:\Users\pnerseth\Amazon Drive\Documents\Private\wazalo\regnskap\Oberlo Orders 2017-01-01-2017-12-04.xlsx");
            foreach (var oberloOrder in oberloOrders)
            {
                Console.WriteLine("{0}", oberloOrder);
            }

            var skandiabankenTransactions = Skandiabanken.ReadTransactions(@"C:\Users\pnerseth\Amazon Drive\Documents\Private\wazalo\regnskap\97132735232_2017_01_01-2017_12_10.xlsx");
            foreach (var skandiabankenTransaction in skandiabankenTransactions)
            {
                Console.WriteLine("{0}", skandiabankenTransaction);
            }

            var aliExpressOrders = AliExpress.ReadOrders(@"C:\Users\pnerseth\Amazon Drive\Documents\Private\wazalo\regnskap\AliExpressOrders-2017-12-10_00-59.csv");
            foreach (var aliExpressOrder in aliExpressOrders)
            {
                Console.WriteLine("{0}", aliExpressOrder);
            }

            // get shopify configuration parameters
            string shopifyDomain = ConfigurationManager.AppSettings["ShopifyDomain"];
            string shopifyAPIKey = ConfigurationManager.AppSettings["ShopifyAPIKey"];
            string shopifyAPIPassword = ConfigurationManager.AppSettings["ShopifyAPIPassword"];

            var shopifyOrders = Shopify.ReadShopifyOrders(shopifyDomain, shopifyAPIKey, shopifyAPIPassword);
            foreach (var order in shopifyOrders)
            {
                Console.WriteLine(order);
            }

            Console.ReadLine();
        }

    }
}
