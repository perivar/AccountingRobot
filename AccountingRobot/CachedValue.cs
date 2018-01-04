using CsvHelper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;

namespace AccountingRobot
{
    public abstract class CachedValue<T>
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
                from = lastCacheFileInfo.To;
                to = currentDate;

                // if the from date is today, then we already have an updated file so use cache
                if (from.Date.Equals(to.Date))
                {
                    // use latest cache file (or force an update)
                    return GetLatest(lastCacheFileInfo.FilePath, from, to, forceUpdate);
                }
                else if (from != firstDayOfTheYear)
                {
                    // combine new and old values
                    var updatedValues = GetCombinedUpdatedAndExisting(lastCacheFileInfo, from, to);

                    // and store to new file
                    string newCacheFilePath = Path.Combine(cacheDir, string.Format("{0}-{1:yyyy-MM-dd}-{2:yyyy-MM-dd}.csv", CacheFileNamePrefix, firstDayOfTheYear, to));
                    using (var sw = new StreamWriter(newCacheFilePath))
                    {
                        var csvWriter = new CsvWriter(sw);
                        csvWriter.Configuration.Delimiter = ",";
                        csvWriter.Configuration.HasHeaderRecord = true;
                        csvWriter.Configuration.CultureInfo = CultureInfo.InvariantCulture;

                        csvWriter.WriteRecords(updatedValues);
                    }

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

            // get updated transactions (or from cache file if update is forced)
            string cacheFilePath = Path.Combine(cacheDir, string.Format("{0}-{1:yyyy-MM-dd}-{2:yyyy-MM-dd}.csv", CacheFileNamePrefix, from, to));
            return GetLatest(cacheFilePath, from, to, forceUpdate);
        }

        public abstract List<T> GetLatest(string filePath, DateTime from, DateTime to, bool forceUpdate);

        public abstract List<T> GetCombinedUpdatedAndExisting(FileDate lastCacheFileInfo, DateTime from, DateTime to);
    }
}
