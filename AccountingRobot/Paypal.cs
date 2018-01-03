using PayPal.PayPalAPIInterfaceService.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using PayPal.PayPalAPIInterfaceService;
using System.Configuration;
using System.IO;
using CsvHelper;
using System.Globalization;

namespace AccountingRobot
{
    public static class Paypal
    {
        public static List<PaymentTransactionSearchResultType> GetTransactions()
        {
            // get paypal configuration parameters
            string payPalApiUsername = ConfigurationManager.AppSettings["PayPalApiUsername"];
            string payPalApiPassword = ConfigurationManager.AppSettings["PayPalApiPassword"];
            string payPalApiSignature = ConfigurationManager.AppSettings["PayPalApiSignature"];

            DateTime endDate = new DateTime(2017, 1, 1);
            DateTime startDate = new DateTime(2017, 12, 31);

            TransactionSearchReq req = new TransactionSearchReq();
            req.TransactionSearchRequest = new TransactionSearchRequestType();

            //req.TransactionSearchRequest.StartDate = startDate.ToString("yyyy-MM-ddTHH:mm:ssZ");
            //req.TransactionSearchRequest.EndDate = endDate.ToString("yyyy-MM-ddTHH:mm:ssZ");

            req.TransactionSearchRequest.StartDate = "2017-01-01T00:00:00Z";
            req.TransactionSearchRequest.EndDate = "2017-12-31T00:00:00Z";

            Dictionary<string, string> config = new Dictionary<string, string>();
            config.Add("mode", "live");
            config.Add("account1.apiUsername", payPalApiUsername);
            config.Add("account1.apiPassword", payPalApiPassword);
            config.Add("account1.apiSignature", payPalApiSignature);

            PayPalAPIInterfaceServiceService service = new PayPalAPIInterfaceServiceService(config);

            TransactionSearchResponseType transactionSearchResponseType = service.TransactionSearch(req);
            if (transactionSearchResponseType.Ack == AckCodeType.FAILURE)
            {
                foreach (var e in transactionSearchResponseType.Errors)
                {
                    Console.WriteLine(e.LongMessage);
                }
            }

            return transactionSearchResponseType.PaymentTransactions;
        }

        public static List<PayPalTransaction> GetLatestPaypalTransactions(bool forceUpdate = false)
        {
            // get paypal configuration parameters
            string payPalApiUsername = ConfigurationManager.AppSettings["PayPalApiUsername"];
            string payPalApiPassword = ConfigurationManager.AppSettings["PayPalApiPassword"];
            string payPalApiSignature = ConfigurationManager.AppSettings["PayPalApiSignature"];
            string cacheDir = ConfigurationManager.AppSettings["CacheDir"];
            string cacheFileNamePrefix = "PayPal Transactions";

            var lastCacheFile = Utils.FindLastCacheFile(cacheDir, cacheFileNamePrefix);

            var date = new Date();
            var currentDate = date.CurrentDate;
            var firstDayOfTheYear = date.FirstDayOfTheYear;

            // check if we have a cache file
            DateTime from = default(DateTime);
            DateTime to = default(DateTime);

            // if the cache file object has values
            if (!lastCacheFile.Equals(default(KeyValuePair<DateTime, string>)))
            {
                from = lastCacheFile.Key.Date;
                to = currentDate;

                // if the from date is today, then we already have an updated file so use cache
                if (from.Date.Equals(to.Date))
                {
                    // use latest cache file (or force an update)
                    return GetPayPalTransactions(lastCacheFile.Value, payPalApiUsername, payPalApiPassword, payPalApiSignature, from, to, forceUpdate);
                }
                else if (from != firstDayOfTheYear)
                {
                    // we have to combine two files:
                    // the original cache file and the new transactions file
                    Console.Out.WriteLine("Finding PayPal transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
                    var newPayPalTransactions = GetPayPalTransactions(payPalApiUsername, payPalApiPassword, payPalApiSignature, from, to);
                    var originalPayPalTransactions = Utils.ReadCacheFile<PayPalTransaction>(lastCacheFile.Value);

                    // copy all the original PayPal transactions into a new file, except entries that are 
                    // from the from date or newer
                    var updatedPayPalTransactions = originalPayPalTransactions.Where(p => p.Timestamp < from).ToList();

                    // and add the new transactions to beginning of list
                    updatedPayPalTransactions.InsertRange(0, newPayPalTransactions);

                    // and store to new file
                    string newCacheFilePath = Path.Combine(cacheDir, string.Format("{0}-{1:yyyy-MM-dd}-{2:yyyy-MM-dd}.csv", cacheFileNamePrefix, firstDayOfTheYear, to));
                    using (var sw = new StreamWriter(newCacheFilePath))
                    {
                        var csvWriter = new CsvWriter(sw);
                        csvWriter.Configuration.Delimiter = ",";
                        csvWriter.Configuration.HasHeaderRecord = true;
                        csvWriter.Configuration.CultureInfo = CultureInfo.InvariantCulture;

                        csvWriter.WriteRecords(updatedPayPalTransactions);
                    }

                    Console.Out.WriteLine("Successfully wrote file to {0}", newCacheFilePath);
                    return updatedPayPalTransactions;
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
            return GetPayPalTransactions(cacheFilePath, payPalApiUsername, payPalApiPassword, payPalApiSignature, from, to, false);
        }

        static List<PayPalTransaction> GetPayPalTransactions(string cacheFilePath, string payPalApiUsername, string payPalApiPassword, string payPalApiSignature, DateTime from, DateTime to, bool forceUpdate = false)
        {
            var cachedPayPalTransactions = Utils.ReadCacheFile<PayPalTransaction>(cacheFilePath, forceUpdate);
            //if (cachedPayPalTransactions != null && cachedPayPalTransactions.Count() > 0)
            if (cachedPayPalTransactions != null)
            {
                Console.Out.WriteLine("Using cache file {0}.", cacheFilePath);
                return cachedPayPalTransactions;
            }
            else
            {
                Console.Out.WriteLine("Finding PayPal transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
                var payPalTransactions = GetPayPalTransactions(payPalApiUsername, payPalApiPassword, payPalApiSignature, from, to);

                using (var sw = new StreamWriter(cacheFilePath))
                {
                    var csvWriter = new CsvWriter(sw);
                    csvWriter.Configuration.Delimiter = ",";
                    csvWriter.Configuration.HasHeaderRecord = true;
                    csvWriter.Configuration.CultureInfo = CultureInfo.InvariantCulture;

                    csvWriter.WriteRecords(payPalTransactions);
                }

                Console.Out.WriteLine("Successfully wrote file to {0}", cacheFilePath);
                return payPalTransactions;
            }
        }

        static List<PayPalTransaction> GetPayPalTransactions(string payPalApiUsername, string payPalApiPassword, string payPalApiSignature, DateTime from, DateTime to)
        {
            TransactionSearchReq req = new TransactionSearchReq();
            req.TransactionSearchRequest = new TransactionSearchRequestType();

            //req.TransactionSearchRequest.StartDate = "2017-01-01T00:00:00Z";
            //req.TransactionSearchRequest.EndDate = "2017-12-31T00:00:00Z";

            string startDate = string.Format("{0:yyyy-MM-ddTHH\\:mm\\:ssZ}", from);
            string endDate = string.Format("{0:yyyy-MM-ddTHH\\:mm\\:ssZ}", to.AddDays(1));
            req.TransactionSearchRequest.StartDate = startDate;
            req.TransactionSearchRequest.EndDate = endDate;

            Dictionary<string, string> config = new Dictionary<string, string>();
            config.Add("mode", "live");
            config.Add("account1.apiUsername", payPalApiUsername);
            config.Add("account1.apiPassword", payPalApiPassword);
            config.Add("account1.apiSignature", payPalApiSignature);

            PayPalAPIInterfaceServiceService service = new PayPalAPIInterfaceServiceService(config);

            TransactionSearchResponseType transactionSearchResponseType = service.TransactionSearch(req);
            if (transactionSearchResponseType.Ack == AckCodeType.FAILURE)
            {
                foreach (var e in transactionSearchResponseType.Errors)
                {
                    Console.WriteLine(e.LongMessage);
                }
            }

            var payPalTransactions = new List<PayPalTransaction>();
            foreach (var transaction in transactionSearchResponseType.PaymentTransactions)
            {
                var payPalTransaction = new PayPalTransaction();

                payPalTransaction.TransactionID = transaction.TransactionID;

                // Converting from paypal date to date:
                // 2017-08-30T21:13:37Z
                // var date = DateTimeOffset.Parse(paypalTransaction.Timestamp).UtcDateTime;
                payPalTransaction.Timestamp = DateTimeOffset.Parse(transaction.Timestamp).UtcDateTime;

                payPalTransaction.Status = transaction.Status;
                payPalTransaction.Type = transaction.Type;

                payPalTransaction.GrossAmount = decimal.Parse(transaction.GrossAmount.value, CultureInfo.InvariantCulture);
                payPalTransaction.NetAmount = decimal.Parse(transaction.NetAmount.value, CultureInfo.InvariantCulture);
                payPalTransaction.FeeAmount = decimal.Parse(transaction.FeeAmount.value, CultureInfo.InvariantCulture);

                if (null != transaction.Payer)
                {
                    payPalTransaction.Payer = transaction.Payer;
                }
                payPalTransaction.PayerDisplayName = transaction.PayerDisplayName;

                payPalTransactions.Add(payPalTransaction);
            }
            return payPalTransactions;
        }
    }
}
