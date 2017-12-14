using System;
using System.Configuration;

namespace AccountingRobot
{
    partial class Program
    {
        static void Main(string[] args)
        {
            // get paypal configuration parameters
            string payPalApiUsername = ConfigurationManager.AppSettings["PayPalApiUsername"];
            string payPalApiPassword = ConfigurationManager.AppSettings["PayPalApiPassword"];
            string payPalApiSignature = ConfigurationManager.AppSettings["PayPalApiSignature"];
            var paypalTransactions = Paypal.GetTransactions(payPalApiUsername, payPalApiPassword, payPalApiSignature);
            foreach (var paypalTransaction in paypalTransactions)
            {
                Console.WriteLine("{0} {1} {2} {3}", paypalTransaction.Timestamp, paypalTransaction.GrossAmount.value, paypalTransaction.FeeAmount.value, paypalTransaction.PayerDisplayName);
            }

            /*
            // get stripe configuration parameters
            string stripeApiKey = ConfigurationManager.AppSettings["StripeApiKey"];
            var stripeTransactions = Stripe.GetTransactions(stripeApiKey);
            foreach (var stripeTransaction in stripeTransactions)
            {
                decimal amount = (decimal)stripeTransaction.Amount / 100;
                decimal fee = (decimal)stripeTransaction.Fee / 100;
                Console.WriteLine("{0:yyyy.MM.dd} {1:N} {2:N} {3} {4} {5} {6}", stripeTransaction.Created, amount, fee, stripeTransaction.Currency, stripeTransaction.Description, stripeTransaction.Type, stripeTransaction.Status);
            }

            
            var oberloOrders = Oberlo.ReadOrders(@"C:\Users\pnerseth\Amazon Drive\Documents\Private\wazalo\regnskap\Oberlo Orders 2017-01-01-2017-12-04.xlsx");
            foreach (var oberloOrder in oberloOrders)
            {
                Console.WriteLine("{0}", oberloOrder);
            }
            

            var skandiabankenTransactions = Skandiabanken.ReadTransactions(@"C:\Users\pnerseth\Amazon Drive\Documents\Private\wazalo\regnskap\97132735232_2017_01_01-2017_12_10.xlsx");
            foreach (var skandiabankenTransaction in skandiabankenTransactions)
            {
                Console.WriteLine("{0}", skandiabankenTransaction.GuessAccountType());
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
            */

            Console.ReadLine();
        }

    }
}
