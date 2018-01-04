using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;

namespace AccountingRobot
{
    public class PayPalFactory : CachedList<PayPalTransaction>
    {
        // get paypal configuration parameters
        string payPalApiUsername = ConfigurationManager.AppSettings["PayPalApiUsername"];
        string payPalApiPassword = ConfigurationManager.AppSettings["PayPalApiPassword"];
        string payPalApiSignature = ConfigurationManager.AppSettings["PayPalApiSignature"];

        public static readonly PayPalFactory Instance = new PayPalFactory();

        private PayPalFactory()
        {
        }

        protected override string CacheFileNamePrefix { get { return "PayPal Transactions"; } }

        public override List<PayPalTransaction> GetCombinedUpdatedAndExisting(FileDate lastCacheFileInfo, DateTime from, DateTime to)
        {
            // we have to combine two files:
            // the original cache file and the new transactions file
            Console.Out.WriteLine("Finding PayPal transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
            var newPayPalTransactions = Paypal.GetPayPalTransactions(payPalApiUsername, payPalApiPassword, payPalApiSignature, from, to);
            var originalPayPalTransactions = Utils.ReadCacheFile<PayPalTransaction>(lastCacheFileInfo.FilePath);

            // copy all the original PayPal transactions into a new file, except entries that are 
            // from the from date or newer
            var updatedPayPalTransactions = originalPayPalTransactions.Where(p => p.Timestamp < from).ToList();

            // and add the new transactions to beginning of list
            updatedPayPalTransactions.InsertRange(0, newPayPalTransactions);

            return updatedPayPalTransactions;
        }

        public override List<PayPalTransaction> GetList(DateTime from, DateTime to)
        {
            Console.Out.WriteLine("Finding PayPal transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
            return Paypal.GetPayPalTransactions(payPalApiUsername, payPalApiPassword, payPalApiSignature, from, to);
        }
    }
}
