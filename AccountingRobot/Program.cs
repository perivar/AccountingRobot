using CsvHelper;
using CsvHelper.TypeConversion;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;

namespace AccountingRobot
{
    partial class Program
    {
        static void Main(string[] args)
        {
            var accountingItems = ProcessAll();
            accountingItems.Reverse();
            var now = DateTime.Now;
            var fileName = string.Format("Accounting {0:yyyy-MM-dd}.csv", now);
            using (var sw = new StreamWriter(fileName))
            {
                var csvWriter = new CsvWriter(sw);
                csvWriter.Configuration.Delimiter = ",";
                csvWriter.Configuration.HasHeaderRecord = true;
                csvWriter.Configuration.CultureInfo = CultureInfo.InvariantCulture;
                csvWriter.Configuration.RegisterClassMap<AccountingItemCsvMap>();
                csvWriter.WriteRecords(accountingItems);
            }
            /*
            // get paypal configuration parameters
            string payPalApiUsername = ConfigurationManager.AppSettings["PayPalApiUsername"];
            string payPalApiPassword = ConfigurationManager.AppSettings["PayPalApiPassword"];
            string payPalApiSignature = ConfigurationManager.AppSettings["PayPalApiSignature"];
            var paypalTransactions = Paypal.GetTransactions(payPalApiUsername, payPalApiPassword, payPalApiSignature);
            foreach (var paypalTransaction in paypalTransactions)
            {
                // 2017-08-30T21:13:37Z
                var date = DateTimeOffset.Parse(paypalTransaction.Timestamp).UtcDateTime;

                Console.WriteLine("{0:dd.MM.yyyy} {1} {2} {3}", date, paypalTransaction.GrossAmount.value, paypalTransaction.FeeAmount.value, paypalTransaction.PayerDisplayName);
            }
            */

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

        static List<AccountingItem> ProcessAll()
        {
            var accountingList = new List<AccountingItem>();

            string skandiabankenXLSX = @"C:\Users\pnerseth\Amazon Drive\Documents\Private\wazalo\regnskap\97132735232_2017_01_01-2017_12_10.xlsx";
            string aliExpressCSV = @"C:\Users\pnerseth\Amazon Drive\Documents\Private\wazalo\regnskap\AliExpressOrders-2017-12-10_00-59.csv";
            //string oberloXLSX = @"C:\Users\pnerseth\Amazon Drive\Documents\Private\wazalo\regnskap\Oberlo Orders 2017-01-01-2017-12-04.xlsx";
            string oberloCSV = @"C:\Users\pnerseth\Amazon Drive\Documents\Private\wazalo\regnskap\Oberlo Orders 2017-01-01-2017-12-31.csv";

            // prepopulate some lookup lists
            var aliExpressOrders = AliExpress.ReadOrders(aliExpressCSV);
            var aliExpressOrderGroups = AliExpress.CombineOrders(aliExpressOrders);
            //var oberloOrders = Oberlo.ReadOrders(oberloXLSX);
            var oberloOrders = Oberlo.ReadOrdersV2(oberloCSV);

            // run through the bank account transactions
            var skandiabankenTransactions = Skandiabanken.ReadTransactions(skandiabankenXLSX);

            // and map each one to the right meta information
            foreach (var skandiabankenTransaction in skandiabankenTransactions)
            {
                var accountingItem = new AccountingItem();

                // extract properties from the transaction text
                skandiabankenTransaction.ExtractAccountingInformation();
                var accountingType = skandiabankenTransaction.AccountingType;

                // 1. If AliExpress purchase
                if (accountingType == SkandiabankenTransaction.AccountingTypeEnum.CostOfGoods)
                {
                    Console.WriteLine("{0}", skandiabankenTransaction);
                    accountingItem.Date = skandiabankenTransaction.TransactionDate;
                    accountingItem.Text = string.Format("{0} {1:dd.MM} {2} {3} {4}", skandiabankenTransaction.GetAccountingTypeString(), skandiabankenTransaction.ExternalPurchaseDate, skandiabankenTransaction.ExternalPurchaseVendor, skandiabankenTransaction.ExternalPurchaseAmount, skandiabankenTransaction.ExternalPurchaseCurrency);
                    accountingItem.NumPurchase = "";
                    accountingItem.PurchaseUSD = skandiabankenTransaction.ExternalPurchaseAmount;
                    accountingItem.AccountBank = skandiabankenTransaction.AccountChange;
                    accountingItem.CostOfGoods = -skandiabankenTransaction.AccountChange;

                    // lookup in AliExpress purchase list
                    // matching ordertime and orderamount
                    var aliExpressQuery =
                        from order in aliExpressOrderGroups
                        where
                        order.OrderTime.Date == skandiabankenTransaction.ExternalPurchaseDate.Date &&
                        order.OrderAmount == skandiabankenTransaction.ExternalPurchaseAmount
                        orderby order.OrderTime ascending
                        select order;

                    if (aliExpressQuery.Count() > 1)
                    {
                        string aliexOrders = String.Join("\n\t", aliExpressQuery.Select(o => o.ToString()));
                        Console.WriteLine("\tERROR: MUST CHOOSE ONE OF MULTIPLE:\n\t{0}", aliexOrders);

                        // flatten the aliexpress order list
                        var aliExpressOrderList = aliExpressQuery.SelectMany(a => a.Children.ToList()).ToList();

                        // join the aliexpress list and the oberlo list on aliexpress order number
                        var joined = from a in aliExpressOrderList
                                     join b in oberloOrders
                                    on a.OrderId.ToString() equals b.AliOrderNumber
                                     select new { AliExpress = a, Oberlo = b };

                        if (joined.Count() > 0)
                        {
                            // found shopify order numbers
                            Console.WriteLine("\tFOUND SHOPIFY ORDERS:");
                            foreach (var join in joined)
                            {
                                Console.WriteLine("\t{0} {1}", join.Oberlo, join.AliExpress);
                                accountingItem.Text += string.Format(" {0},", join.Oberlo.OrderNumber);
                                accountingItem.NumPurchase += string.Format("{0}, ", join.Oberlo.OrderNumber);
                            }
                        } else
                        {
                            // could not find shopify order numbers
                            Console.WriteLine("\tERROR: NO SHOPIFY ORDERS FOUND!");
                        }
                    }
                    else if (aliExpressQuery.Count() == 1)
                    {
                        Console.WriteLine("\tOK: FOUND SINGLE: {0}", aliExpressQuery.First());

                        // join order ids and make sure they are strings
                        var idsLong = aliExpressQuery.SelectMany(a => a.Children.ToList().Select(b => b.OrderId)).ToList();
                        var idsString = idsLong.ConvertAll<string>(x => x.ToString());

                        // lookup in oberlo to find shopify order number
                        var oberloQuery =
                            from order in oberloOrders
                            where
                            idsString.Contains(order.AliOrderNumber)
                            orderby order.CreatedDate ascending
                            select order;

                        foreach (var oberlo in oberloQuery)
                        {
                            Console.WriteLine("\t{0}", oberlo);
                            accountingItem.Text += string.Format(" {0},", oberlo.OrderNumber);
                            accountingItem.NumPurchase += string.Format("{0}, ", oberlo.OrderNumber);
                        }
                    }
                    else
                    {
                        Console.WriteLine("\tERROR: NONE FOUND!");
                    }
                }

                // 2. Transfer Paypal
                else if (accountingType == SkandiabankenTransaction.AccountingTypeEnum.TransferPaypal)
                {
                    Console.WriteLine("{0}", skandiabankenTransaction);

                    accountingItem.Date = skandiabankenTransaction.TransactionDate;
                    accountingItem.Text = string.Format("{0} {1:dd.MM} {2}", skandiabankenTransaction.GetAccountingTypeString(), skandiabankenTransaction.ExternalPurchaseDate, skandiabankenTransaction.ExternalPurchaseVendor);
                    accountingItem.AmountNOK = 0; // what it was before the fees?
                    accountingItem.Gateway = "paypal";

                    accountingItem.AccountPaypal = -skandiabankenTransaction.AccountChange;
                    accountingItem.AccountBank = skandiabankenTransaction.AccountChange;
                }

                // 3. Transfer Stripe
                else if (accountingType == SkandiabankenTransaction.AccountingTypeEnum.TransferStripe)
                {
                    Console.WriteLine("{0}", skandiabankenTransaction);

                    accountingItem.Date = skandiabankenTransaction.TransactionDate;
                    accountingItem.Text = string.Format("{0} {1:dd.MM} {2}", skandiabankenTransaction.GetAccountingTypeString(), skandiabankenTransaction.ExternalPurchaseDate, skandiabankenTransaction.ExternalPurchaseVendor);
                    accountingItem.AmountNOK = 0; // what it was before the fees?
                    accountingItem.Gateway = "stripe";

                    accountingItem.AccountStripe = -skandiabankenTransaction.AccountChange;
                    accountingItem.AccountBank = skandiabankenTransaction.AccountChange;
                }

                // 4. None of those above
                else
                {
                    Console.WriteLine("{0}", skandiabankenTransaction);

                    accountingItem.Date = skandiabankenTransaction.TransactionDate;
                    accountingItem.Text = string.Format("{0} {1} {2} {3}", skandiabankenTransaction.GetAccountingTypeString(), skandiabankenTransaction.ArchiveReference, skandiabankenTransaction.Type, skandiabankenTransaction.Text);
                    accountingItem.AccountBank = skandiabankenTransaction.AccountChange;
                }

                accountingList.Add(accountingItem);
            }
            return accountingList;
        }
    }
}
