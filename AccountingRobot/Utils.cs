using CsvHelper;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace AccountingRobot
{
    public static class Utils
    { 
        public static bool CaseInsensitiveContains(this string text, string value, StringComparison stringComparison = StringComparison.CurrentCultureIgnoreCase)
        {
            return text.IndexOf(value, stringComparison) >= 0;
        }

        public static KeyValuePair<DateTime, string> FindLastCacheFile(string cacheDir, string cacheFileNamePrefix)
        {
            string dateFromToRegexPattern = @"(\d{4}\-\d{2}\-\d{2})\-(\d{4}\-\d{2}\-\d{2})\.csv$";
            return FindLastCacheFile(cacheDir, cacheFileNamePrefix, dateFromToRegexPattern, "yyyy-MM-dd", "\\-");
        }

        public static KeyValuePair<DateTime, string> FindLastCacheFile(string cacheDir, string cacheFileNamePrefix, string dateFromToRegexPattern, string dateParsePattern, string separator)
        {
            var dateDictonary = new SortedDictionary<DateTime, string>();

            string regexp = string.Format("{0}{1}{2}", cacheFileNamePrefix, separator, dateFromToRegexPattern);
            Regex reg = new Regex(regexp);

            string directorySearchPattern = string.Format("{0}*", cacheFileNamePrefix);
            IEnumerable<string> filePaths = Directory.EnumerateFiles(cacheDir, directorySearchPattern);
            foreach (var filePath in filePaths)
            {
                var fileName = Path.GetFileName(filePath);
                var match = reg.Match(fileName);
                if (match.Success)
                {
                    var from = match.Groups[1].Value;
                    var to = match.Groups[2].Value;

                    var dateTo = DateTime.ParseExact(to, dateParsePattern, CultureInfo.InvariantCulture);
                    dateDictonary.Add(dateTo, filePath);
                }
            }

            if (dateDictonary.Count() > 0)
            {
                // the first element is the newest date
                return dateDictonary.Last();
            }

            // return a default key value pair
            return default(KeyValuePair<DateTime, string>);
        }

        public static List<T> ReadCacheFile<T>(string filePath, bool forceUpdate = false)
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

                        return csvReader.GetRecords<T>().ToList();
                    }
                }
            }
            else
            {
                return null;
            }
        }
    }

    public class Date
    {
        DateTime currentDate;
        DateTime yesterday;
        DateTime firstDayOfTheYear;
        DateTime lastDayOfTheYear;

        public DateTime CurrentDate {
            get { return currentDate; }
        }

        public DateTime Yesterday
        {
            get { return yesterday; }
        }

        public DateTime FirstDayOfTheYear
        {
            get { return firstDayOfTheYear; }
        }

        public DateTime LastDayOfTheYear
        {
            get { return lastDayOfTheYear; }
        }

        public Date()
        {
            currentDate = DateTime.Now.Date;
            yesterday = currentDate.AddDays(-1);
            firstDayOfTheYear = new DateTime(currentDate.Year, 1, 1);
            lastDayOfTheYear = new DateTime(currentDate.Year, 12, 31);
        }
    }
}
