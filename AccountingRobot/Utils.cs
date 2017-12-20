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

        public static KeyValuePair<DateTime, string> FindLastCacheFile(string directoryPath, string filePrefix)
        {
            var dateDictonary = new SortedDictionary<DateTime, string>();

            string dateFromToRegex = @"(\d{4}\-\d{2}\-\d{2})\-(\d{4}\-\d{2}\-\d{2})\.csv$";
            string regexp = string.Format("{0}\\-{1}", filePrefix, dateFromToRegex);
            Regex reg = new Regex(regexp);

            string directorySearchPattern = string.Format("{0}*", filePrefix);
            IEnumerable<string> filePaths = Directory.EnumerateFiles(directoryPath, directorySearchPattern);
            foreach (var filePath in filePaths)
            {
                var fileName = Path.GetFileName(filePath);
                var match = reg.Match(fileName);
                if (match.Success)
                {
                    var from = match.Groups[1].Value;
                    var to = match.Groups[2].Value;

                    var dateTo = DateTime.ParseExact(to, "yyyy-MM-dd", CultureInfo.InvariantCulture);
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
}
