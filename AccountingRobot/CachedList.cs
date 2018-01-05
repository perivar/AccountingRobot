using CsvHelper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;

namespace AccountingRobot
{
    public abstract class CachedList<T>
    {
        protected abstract string CacheFileNamePrefix { get; }

        public List<T> GetLatest(bool forceUpdate = false)
        {
            string cacheDir = ConfigurationManager.AppSettings["CacheDir"];

            var lastCacheFileInfo = Utils.FindLastCacheFile(cacheDir, CacheFileNamePrefix);

            var date = new Date();
            var currentDate = date.CurrentDate;
            var firstDayOfTheYear = date.FirstDayOfTheYear;

            // check if we have a cache file
            DateTime from = default(DateTime);
            DateTime to = default(DateTime);

            // if the cache file object has values
            if (!lastCacheFileInfo.Equals(default(FileDate)))
            {
                // find values starting from when the cache file ends and until now
                from = lastCacheFileInfo.To;
                to = currentDate;
                
                // if the from date is today, then we already have an updated file so use cache
                if (lastCacheFileInfo.To.Date.Equals(currentDate.Date))
                {
                    // use latest cache file (or force an update)
                    return GetList(lastCacheFileInfo.FilePath, from, to, forceUpdate);
                }
                else if (from != firstDayOfTheYear)
                {
                    // combine new and old values
                    var updatedValues = GetCombinedUpdatedAndExisting(lastCacheFileInfo, from, to);

                    // and store to new file
                    string newCacheFilePath = Path.Combine(cacheDir, string.Format("{0}-{1:yyyy-MM-dd}-{2:yyyy-MM-dd}.csv", CacheFileNamePrefix, firstDayOfTheYear, to));
                    Utils.WriteCacheFile(newCacheFilePath, updatedValues);
                    Console.Out.WriteLine("Successfully wrote file to {0}", newCacheFilePath);
                    return updatedValues;
                }
            }
            else
            {
                // find all from beginning of year until now
                from = firstDayOfTheYear;
                to = currentDate;
            }

            // get updated transactions (or from cache file if update is not forced)
            string cacheFilePath = Path.Combine(cacheDir, string.Format("{0}-{1:yyyy-MM-dd}-{2:yyyy-MM-dd}.csv", CacheFileNamePrefix, from, to));
            return GetList(cacheFilePath, from, to, forceUpdate);
        }

        public List<T> GetList(string cacheFilePath, DateTime from, DateTime to, bool forceUpdate)
        {
            var cachedList = Utils.ReadCacheFile<T>(cacheFilePath, forceUpdate);
            //if (cachedList != null && cachedList.Count() > 0)
            if (cachedList != null)
            {
                Console.Out.WriteLine("Using cache file {0}.", cacheFilePath);
                return cachedList;
            }
            else
            {
                var values = GetList(from, to);
                Utils.WriteCacheFile(cacheFilePath, values);
                Console.Out.WriteLine("Successfully wrote file to {0}", cacheFilePath);
                return values;
            }
        }

        public abstract List<T> GetList(DateTime from, DateTime to);

        public abstract List<T> GetCombinedUpdatedAndExisting(FileDate lastCacheFileInfo, DateTime from, DateTime to);
    }
}
