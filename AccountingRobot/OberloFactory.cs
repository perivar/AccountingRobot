using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using OberloScraper;

namespace AccountingRobot
{
    public class OberloFactory : CachedList<OberloOrder>
    {
        // get oberlo configuration parameters
        string userDataDir = ConfigurationManager.AppSettings["UserDataDir"];
        string oberloUsername = ConfigurationManager.AppSettings["OberloUsername"];
        string oberloPassword = ConfigurationManager.AppSettings["OberloPassword"];

        public static readonly OberloFactory Instance = new OberloFactory();

        private OberloFactory()
        {
        }

        protected override string CacheFileNamePrefix { get { return "Oberlo Orders"; } }

        public override List<OberloOrder> GetCombinedUpdatedAndExisting(FileDate lastCacheFileInfo, DateTime from, DateTime to)
        {
            // we have to combine two files:
            // the original cache file and the new transactions file
            Console.Out.WriteLine("Finding Oberlo Orders from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
            var newOberloOrders = Oberlo.ScrapeOberloOrders(userDataDir, oberloUsername, oberloPassword, from, to);
            var originalOberloOrders = Utils.ReadCacheFile<OberloOrder>(lastCacheFileInfo.FilePath);

            // copy all the original Oberlo orders into a new file, except entries that are 
            // from the from date or newer
            var updatedOberloOrders = originalOberloOrders.Where(p => p.CreatedDate < from).ToList();

            // and add the new orders to beginning of list
            updatedOberloOrders.InsertRange(0, newOberloOrders);

            return updatedOberloOrders;
        }

        public override List<OberloOrder> GetList(DateTime from, DateTime to)
        {
            Console.Out.WriteLine("Finding Oberlo Orders from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
            return Oberlo.ScrapeOberloOrders(userDataDir, oberloUsername, oberloPassword, from, to);
        }
    }
}
