using CsvHelper;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace AccountingRobot
{
    public static class AliExpress
    {
        public static List<AliExpressOrder> ReadOrders(string aliExpressOrdersFilePath)
        {
            using (TextReader fileReader = File.OpenText(aliExpressOrdersFilePath))
            {
                fileReader.ReadLine(); // skip excel separator line

                using (var csvReader = new CsvReader(fileReader))
                {
                    csvReader.Configuration.Delimiter = ",";
                    csvReader.Configuration.HasHeaderRecord = true;
                    csvReader.Configuration.CultureInfo = CultureInfo.InvariantCulture;
                    csvReader.Configuration.RegisterClassMap<CsvMap>();

                    return csvReader.GetRecords<AliExpressOrder>().ToList<AliExpressOrder>();
                }
            }
        }
    }
}
