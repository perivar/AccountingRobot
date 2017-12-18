using CsvHelper;
using OberloScraper;
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
            var accountingShopifyItems = ProcessShopifyStatement();
            var accountingBankItems = ProcessBankAccountStatement();

            // merge into one list
            accountingShopifyItems.AddRange(accountingBankItems);

            // and sort (by ascending)
            var accountingItems = accountingShopifyItems.OrderBy(o => o.Date).ToList();

            // and store in file
            var now = DateTime.Now;
            var fileName = string.Format("Accounting {0:yyyy-MM-dd}.csv", now);
            using (var sw = new StreamWriter(fileName))
            {
                sw.Write("sep=,\n");
                using (var csvWriter = new CsvWriter(sw))
                {
                    csvWriter.Configuration.Delimiter = ",";
                    csvWriter.Configuration.HasHeaderRecord = true;
                    csvWriter.Configuration.CultureInfo = new CultureInfo("no-No");
                    csvWriter.Configuration.RegisterClassMap<AccountingItemCsvMap>();

                    // write all
                    //csvWriter.WriteRecords(accountingItems);

                    // write header
                    csvWriter.WriteHeader<AccountingItem>();
                    csvWriter.NextRecord();

                    // write each record
                    foreach (var accountingItem in accountingItems)
                    {
                        csvWriter.WriteRecord<AccountingItem>(accountingItem);
                        csvWriter.NextRecord();
                    }
                }
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
            
            // get stripe configuration parameters
            string stripeApiKey = ConfigurationManager.AppSettings["StripeApiKey"];
            var stripeTransactions = Stripe.GetCharges(stripeApiKey);
            foreach (var stripeTransaction in stripeTransactions)
            {
                DateTime date = stripeTransaction.Created;
                string email = stripeTransaction.Metadata["email"];
                string orderId = stripeTransaction.Metadata["order_id"];
                string chargeId = stripeTransaction.Id;
                bool paid = stripeTransaction.Paid;
                bool refunded = stripeTransaction.Refunded;
                string status = stripeTransaction.Status;

                decimal amount = 0;
                decimal fee = 0;
                decimal net = 0;
                if (stripeTransaction.BalanceTransaction != null)
                {
                    amount = (decimal)stripeTransaction.BalanceTransaction.Amount / 100;
                    fee = (decimal)stripeTransaction.BalanceTransaction.Fee / 100;
                    net = (decimal)stripeTransaction.BalanceTransaction.Net / 100;
                }
                Console.WriteLine("{0:yyyy.MM.dd} {1} {2} {3:N} {4:N} {5:N} {6} {7} {8} {9}", date, orderId, chargeId, amount, fee, net, email, status, paid ? "paid" : "not paid", refunded ? "refunded" : "not refunded");
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

            */

            Console.ReadLine();
        }

        static List<AccountingItem> ProcessBankAccountStatement()
        {
            var accountingList = new List<AccountingItem>();

            var currentDate = DateTime.Now.Date;
            var currentYear = currentDate.Year;
            var from = new DateTime(currentYear, 1, 1);
            var to = currentDate;

            string userDataDir = ConfigurationManager.AppSettings["UserDataDir"];
            string oberloUsername = ConfigurationManager.AppSettings["OberloUsername"];
            string oberloPassword = ConfigurationManager.AppSettings["OberloPassword"];

            string skandiabankenXLSX = @"C:\Users\pnerseth\Amazon Drive\Documents\Private\wazalo\regnskap\97132735232_2017_01_01-2017_12_15.xlsx";
            string aliExpressCSV = @"C:\Users\pnerseth\Amazon Drive\Documents\Private\wazalo\regnskap\AliExpressOrders-2017-12-18_08-04.csv";
            //string oberloCSV = @"C:\Users\pnerseth\Amazon Drive\Documents\Private\wazalo\regnskap\Oberlo Orders 2017-01-01-2017-12-18.csv";

            // prepopulate some lookup lists
            var oberloOrders = Oberlo.GetLatestOberloOrders(userDataDir, oberloUsername, oberloPassword);
            var aliExpressOrders = AliExpress.ReadOrders(aliExpressCSV);
            var aliExpressOrderGroups = AliExpress.CombineOrders(aliExpressOrders);

            // run through the bank account transactions
            var skandiabankenTransactions = Skandiabanken.ReadTransactions(skandiabankenXLSX);

            // and map each one to the right meta information
            foreach (var skandiabankenTransaction in skandiabankenTransactions)
            {
                // define accounting item
                var accountingItem = new AccountingItem();
                accountingItem.Date = skandiabankenTransaction.TransactionDate;
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
                            } else
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
                        } else
                        {
                            // lookup customer name
                            accountingItem.CustomerName = oberloQuery.First().CustomerName;
                        }
                        Console.WriteLine("\t{0}", orderNumbers);
                        accountingItem.NumPurchase = orderNumbers;
                    }
                    else
                    {
                        Console.WriteLine("\tERROR: NONE FOUND!");
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
                    }
                }

                accountingList.Add(accountingItem);
            }
            return accountingList;
        }

        static List<AccountingItem> ProcessShopifyStatement()
        {
            var accountingList = new List<AccountingItem>();

            // prepopulate some lookup lists
            Console.Out.WriteLine("Prepopulating Lookup Lists ...");
            // get stripe configuration parameters
            string stripeApiKey = ConfigurationManager.AppSettings["StripeApiKey"];
            var stripeTransactions = Stripe.GetCharges(stripeApiKey);
            Console.Out.WriteLine("Successfully read Stripe transactions ...");

            // get paypal configuration parameters
            string payPalApiUsername = ConfigurationManager.AppSettings["PayPalApiUsername"];
            string payPalApiPassword = ConfigurationManager.AppSettings["PayPalApiPassword"];
            string payPalApiSignature = ConfigurationManager.AppSettings["PayPalApiSignature"];
            var paypalTransactions = Paypal.GetTransactions(payPalApiUsername, payPalApiPassword, payPalApiSignature);
            Console.Out.WriteLine("Successfully read PayPal transactions ...");

            // get shopify configuration parameters
            string shopifyDomain = ConfigurationManager.AppSettings["ShopifyDomain"];
            string shopifyAPIKey = ConfigurationManager.AppSettings["ShopifyAPIKey"];
            string shopifyAPIPassword = ConfigurationManager.AppSettings["ShopifyAPIPassword"];

            var shopifyOrders = Shopify.ReadShopifyOrders(shopifyDomain, shopifyAPIKey, shopifyAPIPassword);
            Console.Out.WriteLine("Successfully read all Shopify orders ...");

            Console.Out.WriteLine("Processing started ...");
            foreach (var shopifyOrder in shopifyOrders)
            {
                // || shopifyOrder.Gateway == "Vipps"
                // skip, not paid (pending), cancelled (voided) and fully refunded orders (refunded)
                if (shopifyOrder.FinancialStatus.Equals("refunded")
                    || shopifyOrder.FinancialStatus.Equals("voided")
                    || shopifyOrder.FinancialStatus.Equals("pending")) continue;

                // define accounting item
                var accountingItem = new AccountingItem();
                accountingItem.Date = shopifyOrder.Date;
                accountingItem.ArchiveReference = shopifyOrder.Id;
                accountingItem.Type = string.Format("{0} {1}", shopifyOrder.FinancialStatus, shopifyOrder.FulfillmentStatus);
                accountingItem.AccountingType = "SHOPIFY";
                accountingItem.Text = string.Format("SALG {0} {1}", shopifyOrder.CustomerName, shopifyOrder.PaymentId);
                accountingItem.CustomerName = shopifyOrder.CustomerName;
                if (shopifyOrder.Gateway != null)
                {
                    accountingItem.Gateway = shopifyOrder.Gateway.ToLower();
                }
                accountingItem.NumSale = shopifyOrder.Name;

                var startDate = shopifyOrder.Date.AddDays(-1);
                var endDate = shopifyOrder.Date.AddDays(1);

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
                        transaction.Metadata["email"].Equals(shopifyOrder.CustomerEmail) &&
                        transaction.Amount == (int)(shopifyOrder.TotalPrice * 100) &&
                         (transaction.Created.Date >= startDate.Date && transaction.Created.Date <= endDate.Date)
                        orderby transaction.Created ascending
                        select transaction;

                        if (stripeQuery.Count() > 1)
                        {
                            // more than one ?!
                            Console.Out.WriteLine("ERROR: FOUND MORE THAN ONE!");
                            accountingItem.ErrorMessage = "Stripe: More than one found, choose one";
                        }
                        else if (stripeQuery.Count() > 0)
                        {
                            // one match
                            var stripeTransaction = stripeQuery.First();
                            decimal amount = (decimal)stripeTransaction.Amount / 100;
                            decimal net = (decimal)stripeTransaction.BalanceTransaction.Net / 100;
                            decimal fee = (decimal)stripeTransaction.BalanceTransaction.Fee / 100;

                            accountingItem.FeesStripe = fee;
                            accountingItem.AccountStripe = net;
                        }
                        else
                        {
                            Console.Out.WriteLine("ERROR: NONE FOUND!");
                            accountingItem.ErrorMessage = "Stripe: No transactions found";
                        }

                        break;
                    case "paypal":

                        accountingItem.PurchaseOtherCurrency = shopifyOrder.TotalPrice;
                        accountingItem.OtherCurrency = "NOK";

                        // Converting from paypal date to date:
                        // 2017-08-30T21:13:37Z
                        // var date = DateTimeOffset.Parse(paypalTransaction.Timestamp).UtcDateTime;

                        // lookup the paypal transaction
                        var paypalQuery =
                        from transaction in paypalTransactions
                        let grossAmount = decimal.Parse(transaction.GrossAmount.value, CultureInfo.InvariantCulture)
                        let timestamp = DateTimeOffset.Parse(transaction.Timestamp).UtcDateTime
                        where
                        transaction.Status.Equals("Completed")
                        && (null != transaction.Payer && transaction.Payer.Equals(shopifyOrder.CustomerEmail))
                        && (grossAmount == shopifyOrder.TotalPrice)
                        && (timestamp.Date >= startDate.Date && timestamp.Date <= endDate.Date)
                        orderby timestamp ascending
                        select transaction;

                        if (paypalQuery.Count() > 1)
                        {
                            // more than one ?!
                            Console.Out.WriteLine("ERROR: FOUND MORE THAN ONE!");
                            accountingItem.ErrorMessage = "Paypal: More than one found, choose one";
                        }
                        else if (paypalQuery.Count() > 0)
                        {
                            // one match
                            var paypalTransaction = paypalQuery.First();
                            decimal amount = decimal.Parse(paypalTransaction.GrossAmount.value, CultureInfo.InvariantCulture);
                            decimal net = decimal.Parse(paypalTransaction.NetAmount.value, CultureInfo.InvariantCulture);
                            decimal fee = decimal.Parse(paypalTransaction.FeeAmount.value, CultureInfo.InvariantCulture);

                            accountingItem.FeesPaypal = -fee;
                            accountingItem.AccountPaypal = net;
                        }
                        else
                        {
                            Console.Out.WriteLine("ERROR: NONE FOUND!");
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
