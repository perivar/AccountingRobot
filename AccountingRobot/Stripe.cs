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
        public static List<StripeBalanceTransaction> GetBalanceTransactions()
        {
            // get stripe configuration parameters
            string stripeApiKey = ConfigurationManager.AppSettings["StripeApiKey"];

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

        public static List<StripeCharge> GetCharges()
        {
            // get stripe configuration parameters
            string stripeApiKey = ConfigurationManager.AppSettings["StripeApiKey"];

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

        // Get all transactions with type payout from StripeBalanceService
        public static List<StripeTransaction> GetAllPayoutTransactions()
        {
            var stripeBalanceTransactions = Stripe.GetBalanceTransactions();

            var stripePayoutsQuery = (from balanceTransaction in stripeBalanceTransactions
                                      where balanceTransaction.Type == "payout"
                                      select new StripeTransaction
                                      {
                                          Created = balanceTransaction.AvailableOn,
                                          Paid = balanceTransaction.Status == "available",
                                          Amount = (decimal)balanceTransaction.Amount / 100,
                                          Net = (decimal)balanceTransaction.Net / 100,
                                          Fee = (decimal)balanceTransaction.Fee / 100,
                                          TransactionID = balanceTransaction.Id
                                      }).ToList();

            return stripePayoutsQuery;
        }

        #region Charge Transactions (StripeChargeService)
        public static List<StripeTransaction> GetLatestStripeTransactions(bool forceUpdate = false)
        {
            // get stripe configuration parameters
            string stripeApiKey = ConfigurationManager.AppSettings["StripeApiKey"];
            string cacheDir = ConfigurationManager.AppSettings["CacheDir"];
            string cacheFileNamePrefix = "Stripe Transactions";

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
                    return GetStripeTransactions(lastCacheFile.Value, stripeApiKey, from, to, forceUpdate);
                }
                else if (from != firstDayOfTheYear)
                {
                    // we have to combine two files:
                    // the original cache file and the new transactions file
                    Console.Out.WriteLine("Finding Stripe transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
                    var newStripeTransactions = GetStripeTransactions(stripeApiKey, from, to);
                    var originalStripeTransactions = Utils.ReadCacheFile<StripeTransaction>(lastCacheFile.Value);

                    // copy all the original stripe transactions into a new file, except entries that are 
                    // from the from date or newer
                    var updatedStripeTransactions = originalStripeTransactions.Where(p => p.Created < from).ToList();

                    // and add the new transactions to beginning of list
                    updatedStripeTransactions.InsertRange(0, newStripeTransactions);

                    // and store to new file
                    string newCacheFilePath = Path.Combine(cacheDir, string.Format("{0}-{1:yyyy-MM-dd}-{2:yyyy-MM-dd}.csv", cacheFileNamePrefix, firstDayOfTheYear, to));
                    using (var sw = new StreamWriter(newCacheFilePath))
                    {
                        var csvWriter = new CsvWriter(sw);
                        csvWriter.Configuration.Delimiter = ",";
                        csvWriter.Configuration.HasHeaderRecord = true;
                        csvWriter.Configuration.CultureInfo = CultureInfo.InvariantCulture;

                        csvWriter.WriteRecords(updatedStripeTransactions);
                    }

                    Console.Out.WriteLine("Successfully wrote file to {0}", newCacheFilePath);
                    return updatedStripeTransactions;
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
            return GetStripeTransactions(cacheFilePath, stripeApiKey, from, to, forceUpdate);
        }

        static List<StripeTransaction> GetStripeTransactions(string cacheFilePath, string stripeApiKey, DateTime from, DateTime to, bool forceUpdate = false)
        {
            var cachedStripeTransactions = Utils.ReadCacheFile<StripeTransaction>(cacheFilePath, forceUpdate);
            if (cachedStripeTransactions != null && cachedStripeTransactions.Count() > 0)
            {
                Console.Out.WriteLine("Using cache file {0}.", cacheFilePath);
                return cachedStripeTransactions;
            }
            else
            {
                Console.Out.WriteLine("Finding Stripe transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
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
                if (allCharges.Count() > 0) lastId = charges.LastOrDefault().Id;
            }

            var stripeTransactions = new List<StripeTransaction>();
            foreach (var charge in allCharges)
            {
                // only process the charges that were successfull, aka Paid
                if (charge.Paid)
                {
                    var stripeTransaction = new StripeTransaction();
                    stripeTransaction.TransactionID = charge.Id;
                    stripeTransaction.Created = charge.Created;
                    stripeTransaction.Paid = charge.Paid;
                    stripeTransaction.CustomerEmail = charge.Metadata["email"];
                    stripeTransaction.OrderID = charge.Metadata["order_id"];
                    stripeTransaction.Amount = (decimal)charge.Amount / 100;
                    stripeTransaction.Net = (decimal)charge.BalanceTransaction.Net / 100;
                    stripeTransaction.Fee = (decimal)charge.BalanceTransaction.Fee / 100;
                    stripeTransaction.Currency = charge.Currency;
                    stripeTransaction.Description = charge.Description;
                    stripeTransaction.Status = charge.Status;
                    decimal amountRefunded = (decimal)charge.AmountRefunded / 100;
                    if (amountRefunded > 0)
                    {
                        // if anything has been refunded
                        // don't add if amount refunded and amount is the same (full refund)
                        if (stripeTransaction.Amount - amountRefunded == 0)
                        {
                            continue;
                        } else
                        {
                            stripeTransaction.Amount = stripeTransaction.Amount - amountRefunded;
                            stripeTransaction.Net = stripeTransaction.Net - amountRefunded;
                        }                        
                    } 
                    stripeTransactions.Add(stripeTransaction);
                }
            }
            return stripeTransactions;
        }
        #endregion

        #region Payout Transactions (StripeBalanceService)
        public static List<StripeTransaction> GetLatestStripePayoutTransactions(bool forceUpdate = false)
        {
            // get stripe configuration parameters
            string stripeApiKey = ConfigurationManager.AppSettings["StripeApiKey"];
            string cacheDir = ConfigurationManager.AppSettings["CacheDir"];
            string cacheFileNamePrefix = "Stripe Payout Transactions";

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
                    return GetStripePayoutTransactions(lastCacheFile.Value, stripeApiKey, from, to, forceUpdate);
                }
                else if (from != firstDayOfTheYear)
                {
                    // we have to combine two files:
                    // the original cache file and the new transactions file
                    Console.Out.WriteLine("Finding Stripe payout transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
                    var newStripePayoutTransactions = GetStripePayoutTransactions(stripeApiKey, from, to);
                    var originalStripePayoutTransactions = Utils.ReadCacheFile<StripeTransaction>(lastCacheFile.Value);

                    // copy all the original stripe transactions into a new file, except entries that are 
                    // from the from date or newer
                    var updatedStripePayoutTransactions = originalStripePayoutTransactions.Where(p => p.Created < from).ToList();

                    // and add the new transactions to beginning of list
                    updatedStripePayoutTransactions.InsertRange(0, newStripePayoutTransactions);

                    // and store to new file
                    string newCacheFilePath = Path.Combine(cacheDir, string.Format("{0}-{1:yyyy-MM-dd}-{2:yyyy-MM-dd}.csv", cacheFileNamePrefix, firstDayOfTheYear, to));
                    using (var sw = new StreamWriter(newCacheFilePath))
                    {
                        var csvWriter = new CsvWriter(sw);
                        csvWriter.Configuration.Delimiter = ",";
                        csvWriter.Configuration.HasHeaderRecord = true;
                        csvWriter.Configuration.CultureInfo = CultureInfo.InvariantCulture;

                        csvWriter.WriteRecords(updatedStripePayoutTransactions);
                    }

                    Console.Out.WriteLine("Successfully wrote file to {0}", newCacheFilePath);
                    return updatedStripePayoutTransactions;
                }
            }
            else
            {
                // find all from beginning of year until now
                from = firstDayOfTheYear;
                to = currentDate;
            }

            // get updated payout transactions (or from cache file if update is forced)
            string cacheFilePath = Path.Combine(cacheDir, string.Format("{0}-{1:yyyy-MM-dd}-{2:yyyy-MM-dd}.csv", cacheFileNamePrefix, from, to));
            return GetStripePayoutTransactions(cacheFilePath, stripeApiKey, from, to, forceUpdate);
        }

        static List<StripeTransaction> GetStripePayoutTransactions(string cacheFilePath, string stripeApiKey, DateTime from, DateTime to, bool forceUpdate = false)
        {
            var cachedStripeTransactions = Utils.ReadCacheFile<StripeTransaction>(cacheFilePath, forceUpdate);
            if (cachedStripeTransactions != null && cachedStripeTransactions.Count() > 0)
            {
                Console.Out.WriteLine("Using cache file {0}.", cacheFilePath);
                return cachedStripeTransactions;
            }
            else
            {
                Console.Out.WriteLine("Finding Stripe payout transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
                var stripeTransactions = GetStripePayoutTransactions(stripeApiKey, from, to);

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

        static List<StripeTransaction> GetStripePayoutTransactions(string stripeApiKey, DateTime from, DateTime to)
        {
            StripeConfiguration.SetApiKey(stripeApiKey);

            var balanceService = new StripeBalanceService();
            var allBalanceTransactions = new List<StripeBalanceTransaction>();
            var lastId = String.Empty;

            int MAX_PAGINATION = 100;
            int itemsExpected = MAX_PAGINATION;
            while (itemsExpected == MAX_PAGINATION)
            {
                IEnumerable<StripeBalanceTransaction> charges = null;
                if (String.IsNullOrEmpty(lastId))
                {
                    charges = balanceService.List(
                    new StripeBalanceTransactionListOptions()
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
                    charges = balanceService.List(
                    new StripeBalanceTransactionListOptions()
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

                allBalanceTransactions.AddRange(charges);
                if (allBalanceTransactions.Count() > 0) lastId = charges.LastOrDefault().Id;
            }

            var stripeTransactions = new List<StripeTransaction>();
            foreach (var balanceTransaction in allBalanceTransactions)
            {
                // only process the charges that are payouts
                if (balanceTransaction.Type == "payout")
                {
                    var stripeTransaction = new StripeTransaction();
                    stripeTransaction.TransactionID = balanceTransaction.Id;
                    stripeTransaction.Created = balanceTransaction.Created;
                    stripeTransaction.AvailableOn = balanceTransaction.AvailableOn;
                    stripeTransaction.Paid = (balanceTransaction.Status == "available");
                    stripeTransaction.Amount = (decimal)balanceTransaction.Amount / 100;
                    stripeTransaction.Net = (decimal)balanceTransaction.Net / 100;
                    stripeTransaction.Fee = (decimal)balanceTransaction.Fee / 100;
                    stripeTransaction.Currency = balanceTransaction.Currency;
                    stripeTransaction.Description = balanceTransaction.Description;
                    stripeTransaction.Status = balanceTransaction.Status;

                    stripeTransactions.Add(stripeTransaction);
                }
            }
            return stripeTransactions;
        }
        #endregion
    }
}
