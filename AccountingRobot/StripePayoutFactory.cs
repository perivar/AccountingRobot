using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;

namespace AccountingRobot
{
    public class StripePayoutFactory : CachedValue<StripeTransaction>
    {
        // get stripe configuration parameters
        string stripeApiKey = ConfigurationManager.AppSettings["StripeApiKey"];

        public static readonly StripePayoutFactory Instance = new StripePayoutFactory();

        private StripePayoutFactory()
        {
        }

        protected override string CacheFileNamePrefix { get { return "Stripe Payout Transactions"; } }

        public override List<StripeTransaction> GetCombinedUpdatedAndExisting(FileDate lastCacheFileInfo, DateTime from, DateTime to)
        {
            // we have to combine two files:
            // the original cache file and the new transactions file
            Console.Out.WriteLine("Finding Stripe payout transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
            var newStripePayoutTransactions = Stripe.GetStripePayoutTransactions(stripeApiKey, from, to);
            var originalStripePayoutTransactions = Utils.ReadCacheFile<StripeTransaction>(lastCacheFileInfo.FilePath);

            // copy all the original stripe transactions into a new file, except entries that are 
            // from the from date or newer
            var updatedStripePayoutTransactions = originalStripePayoutTransactions.Where(p => p.Created < from).ToList();

            // and add the new transactions to beginning of list
            updatedStripePayoutTransactions.InsertRange(0, newStripePayoutTransactions);

            return updatedStripePayoutTransactions;
        }

        public override List<StripeTransaction> GetLatest(string filePath, DateTime from, DateTime to, bool forceUpdate)
        {
            return Stripe.GetStripePayoutTransactions(filePath, stripeApiKey, from, to, forceUpdate);
        }
    }
}
