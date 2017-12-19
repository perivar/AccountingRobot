using CsvHelper;
using Stripe;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;

namespace AccountingRobot
{
    public static class Stripe
    {
        public static List<StripeBalanceTransaction> GetBalanceTransactions(string stripeApiKey)
        {
            StripeConfiguration.SetApiKey(stripeApiKey);
            var balanceService = new StripeBalanceService();
            var allBalanceTransactions = new List<StripeBalanceTransaction>();

            var lastId = String.Empty;

            int MAX_PAGINATION = 100;
            int itemsExpected = MAX_PAGINATION;
            while (itemsExpected == MAX_PAGINATION)
            {
                IEnumerable<StripeBalanceTransaction> balanceTransactions = null;
                if (String.IsNullOrEmpty(lastId))
                {
                    balanceTransactions = balanceService.List(
                    new StripeBalanceTransactionListOptions() { Limit = MAX_PAGINATION });
                    itemsExpected = balanceTransactions.Count();
                }
                else
                {
                    balanceTransactions = balanceService.List(
                    new StripeBalanceTransactionListOptions()
                    {
                        Limit = MAX_PAGINATION,
                        StartingAfter = lastId
                    });
                    itemsExpected = balanceTransactions.Count();
                }

                allBalanceTransactions.AddRange(balanceTransactions);
                lastId = balanceTransactions.LastOrDefault().Id;
            }

            return allBalanceTransactions;
        }

        public static List<StripeCharge> GetCharges(string stripeApiKey)
        {
            StripeConfiguration.SetApiKey(stripeApiKey);

            var chargeService = new StripeChargeService();
            chargeService.ExpandBalanceTransaction = true;
            chargeService.ExpandCustomer = true;
            chargeService.ExpandInvoice = true;

            var allCharges = new List<StripeCharge>();
            var lastId = String.Empty;

            int MAX_PAGINATION = 100;
            int itemsExpected = MAX_PAGINATION;
            while (itemsExpected == MAX_PAGINATION)
            {
                IEnumerable<StripeCharge> charges = null;
                if (String.IsNullOrEmpty(lastId))
                {
                    charges = chargeService.List(
                    new StripeChargeListOptions() { Limit = MAX_PAGINATION });
                    itemsExpected = charges.Count();
                }
                else
                {
                    charges = chargeService.List(
                    new StripeChargeListOptions()
                    {
                        Limit = MAX_PAGINATION,
                        StartingAfter = lastId
                    });
                    itemsExpected = charges.Count();
                }

                allCharges.AddRange(charges);
                lastId = charges.LastOrDefault().Id;
            }

            return allCharges;
        }

        public static List<StripeTransaction> GetLatestStripeTransactions()
        {
            // get stripe configuration parameters
            string stripeApiKey = ConfigurationManager.AppSettings["StripeApiKey"];
            string cacheDir = ConfigurationManager.AppSettings["CacheDir"];
            string cacheFileNamePrefix = "Stripe Transactions";

            var lastCacheFile = Utils.FindLastCacheFile(cacheDir, cacheFileNamePrefix);

            // check if we have a cache file
            DateTime from = default(DateTime);
            DateTime to = default(DateTime);

            // if the cache file object has values
            if (!lastCacheFile.Equals(default(KeyValuePair<DateTime, string>)))
            {
                var currentDate = DateTime.Now.Date;
                from = lastCacheFile.Key.Date;
                to = currentDate;

                // check that the from date isn't today
                if (from.Equals(to))
                {
                    Console.Out.WriteLine("Latest Stripe cache file is from today.");
                    return GetStripeTransactionsCacheFile(lastCacheFile.Value);
                }
            }
            else
            {
                // find all from beginning of year until now
                var currentDate = DateTime.Now.Date;
                var currentYear = currentDate.Year;
                from = new DateTime(currentYear, 1, 1);
                to = currentDate;
            }

            Console.Out.WriteLine("Finding Stripe transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
            return GetStripeTransactions(cacheDir, cacheFileNamePrefix, stripeApiKey, from, to, false);
        }

        static List<StripeTransaction> GetStripeTransactionsCacheFile(string filePath, bool forceUpdate = false)
        {
            // force update even if cache file exists
            if (forceUpdate) return null;

            if (File.Exists(filePath))
            {
                using (TextReader fileReader = File.OpenText(filePath))
                {
                    using (var csvReader = new CsvReader(fileReader))
                    {
                        csvReader.Configuration.Delimiter = ",";
                        csvReader.Configuration.HasHeaderRecord = true;
                        csvReader.Configuration.CultureInfo = CultureInfo.InvariantCulture;

                        return csvReader.GetRecords<StripeTransaction>().ToList<StripeTransaction>();
                    }
                }
            }
            else
            {
                return null;
            }
        }

        static List<StripeTransaction> GetStripeTransactions(string cacheDir, string cacheFileNamePrefix, string stripeApiKey, DateTime from, DateTime to, bool forceUpdate = false)
        {
            string cacheFilePath = Path.Combine(cacheDir, string.Format("{0}-{1:yyyy-MM-dd}-{2:yyyy-MM-dd}.csv", cacheFileNamePrefix, from, to));

            var cachedStripeTransactions = GetStripeTransactionsCacheFile(cacheFilePath, forceUpdate);
            if (cachedStripeTransactions != null && cachedStripeTransactions.Count() > 0)
            {
                Console.Out.WriteLine("Found cached file.");
                return cachedStripeTransactions;
            }
            else
            {
                var stripeTransactions = GetStripeTransactions(stripeApiKey, from, to);

                using (var sw = new StreamWriter(cacheFilePath))
                {
                    var csvWriter = new CsvWriter(sw);
                    csvWriter.Configuration.Delimiter = ",";
                    csvWriter.Configuration.HasHeaderRecord = true;
                    csvWriter.Configuration.CultureInfo = CultureInfo.InvariantCulture;

                    csvWriter.WriteRecords(stripeTransactions);
                }

                Console.Out.WriteLine("Successfully wrote file to {0}", cacheFilePath);
                return stripeTransactions;
            }
        }

        static List<StripeTransaction> GetStripeTransactions(string stripeApiKey, DateTime from, DateTime to)
        {
            StripeConfiguration.SetApiKey(stripeApiKey);

            var chargeService = new StripeChargeService();
            chargeService.ExpandBalanceTransaction = true;
            chargeService.ExpandCustomer = true;
            chargeService.ExpandInvoice = true;

            var allCharges = new List<StripeCharge>();
            var lastId = String.Empty;

            int MAX_PAGINATION = 100;
            int itemsExpected = MAX_PAGINATION;
            while (itemsExpected == MAX_PAGINATION)
            {
                IEnumerable<StripeCharge> charges = null;
                if (String.IsNullOrEmpty(lastId))
                {
                    charges = chargeService.List(
                    new StripeChargeListOptions()
                    {
                        Limit = MAX_PAGINATION,
                        Created = new StripeDateFilter
                        {
                            GreaterThanOrEqual = from.Date,
                            LessThanOrEqual = to.Date
                        }
                    });
                    itemsExpected = charges.Count();
                }
                else
                {
                    charges = chargeService.List(
                    new StripeChargeListOptions()
                    {
                        Limit = MAX_PAGINATION,
                        StartingAfter = lastId,
                        Created = new StripeDateFilter
                        {
                            GreaterThanOrEqual = from.Date,
                            LessThanOrEqual = to.Date
                        }
                    });
                    itemsExpected = charges.Count();
                }

                allCharges.AddRange(charges);
                lastId = charges.LastOrDefault().Id;
            }

            var stripeTransactions = new List<StripeTransaction>();
            foreach (var charge in allCharges)
            {
                // only process the charges that were successfull, aka Paid
                if (charge.Paid)
                {
                    var stripeTransaction = new StripeTransaction();
                    stripeTransaction.Created = charge.Created;
                    stripeTransaction.Paid = charge.Paid;
                    stripeTransaction.CustomerEmail = charge.Metadata["email"];
                    stripeTransaction.Amount = (decimal)charge.Amount / 100;
                    stripeTransaction.Net = (decimal)charge.BalanceTransaction.Net / 100;
                    stripeTransaction.Fee = (decimal)charge.BalanceTransaction.Fee / 100;
                    stripeTransactions.Add(stripeTransaction);
                }
            }
            return stripeTransactions;
        }
    }
}
