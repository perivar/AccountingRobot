using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using AliOrderScraper;

namespace AccountingRobot
{
    public class AliExpressFactory : CachedList<AliExpressOrder>
    {
        string userDataDir = ConfigurationManager.AppSettings["UserDataDir"];
        string cacheDir = ConfigurationManager.AppSettings["CacheDir"];
        string aliExpressUsername = ConfigurationManager.AppSettings["AliExpressUsername"];
        string aliExpressPassword = ConfigurationManager.AppSettings["AliExpressPassword"];

        public static readonly AliExpressFactory Instance = new AliExpressFactory();

        private AliExpressFactory()
        {
        }

        protected override string CacheFileNamePrefix { get { return "AliExpress Orders"; } }

        public override List<AliExpressOrder> GetCombinedUpdatedAndExisting(FileDate lastCacheFileInfo, DateTime from, DateTime to)
        {
            // we have to combine two files:
            // the original cache file and the new transactions file
            Console.Out.WriteLine("Finding AliExpress Orders from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
            var newAliExpressOrders = AliExpress.ScrapeAliExpressOrders(userDataDir, aliExpressUsername, aliExpressPassword, from);
            var originalAliExpressOrders = Utils.ReadCacheFile<AliExpressOrder>(lastCacheFileInfo.FilePath);

            // copy all the original AliExpress orders into a new file, except entries that are 
            // from the from date or newer
            var updatedAliExpressOrders = originalAliExpressOrders.Where(p => p.OrderTime < from).ToList();

            // and add the new orders to beginning of list
            updatedAliExpressOrders.InsertRange(0, newAliExpressOrders);

            return updatedAliExpressOrders;
        }

        public override List<AliExpressOrder> GetList(DateTime from, DateTime to)
        {
            Console.Out.WriteLine("Finding AliExpress Orders from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
            return AliExpress.ScrapeAliExpressOrders(userDataDir, aliExpressUsername, aliExpressPassword, from);
        }
    }
}
