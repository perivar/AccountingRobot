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
                    csvReader.Configuration.RegisterClassMap<AliExpressCsvMap>();

                    return csvReader.GetRecords<AliExpressOrder>().ToList<AliExpressOrder>();
                }
            }
        }

        public static List<AliExpressOrderGroup> CombineOrders(List<AliExpressOrder> aliExpressOrders)
        {
            var query = (from o in aliExpressOrders
                         group o by new { o.OrderTime, o.ContactName }
             into grp
                         select new AliExpressOrderGroup()
                         {
                             OrderTime = grp.Key.OrderTime,
                             ContactName = grp.Key.ContactName,
                             OrderAmount = grp.Sum(o => o.OrderAmount),
                             Children = grp.ToList()
                         }).ToList();

            return query;
        }
    }

    public class AliExpressOrderGroup
    {
        public DateTime OrderTime { get; set; }
        public decimal OrderAmount { get; set; }
        public string ContactName { get; set; }
        public List<AliExpressOrder> Children { get; set; }

        public override string ToString()
        {
            return string.Format("{0} {1:dd.MM} {2} {3}", string.Join(", ", Array.ConvertAll(Children.ToArray(), i => i.OrderId)), OrderTime, OrderAmount.ToString("C", new CultureInfo("en-US")), ContactName);
        }
    }
}
