using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FastExcel;
using CsvHelper;

namespace AccountingRobot
{
    class Program
    {
        static void Main(string[] args)
        {
            //ReadOberloOrders(@"C:\Users\pnerseth\Amazon Drive\Documents\Private\wazalo\regnskap\Oberlo Orders 2017-01-01-2017-12-04.xlsx");
            ReadAliExpressOrders(@"C:\Users\pnerseth\Amazon Drive\Documents\Private\wazalo\regnskap\AliExpressOrders-2017-12-10_00-59.csv");
            Console.ReadLine();
        }

        static void ReadOberloOrders(string oberloOrdersFilesPath)
        {

            // Get the input file paths
            FileInfo inputFile = new FileInfo(oberloOrdersFilesPath);

            //Create a worksheet
            Worksheet worksheet = null;

            // Create an instance of Fast Excel
            using (var fastExcel = new FastExcel.FastExcel(inputFile, true))
            {
                // Read the rows using worksheet name
                string worksheetName = "Orders";
                worksheet = fastExcel.Read(worksheetName, 1);

                Console.WriteLine("Reading worksheet {0} ...", worksheetName);

                foreach (var row in worksheet.Rows)
                {
                    var orderNumber = row.GetCellByColumnName("A"); // Order #
                    var createdDate = row.GetCellByColumnName("B"); // Created
                    var financialStatus = row.GetCellByColumnName("C"); // Financial status                    
                    var supplier = row.GetCellByColumnName("G"); // Supplier
                    var SKU = row.GetCellByColumnName("H"); // SKU
                    var productName = row.GetCellByColumnName("I"); // Product
                    var variant = row.GetCellByColumnName("J"); // Variant
                    var trackingNumber = row.GetCellByColumnName("L"); // Tracking number
                    var aliOrderNumber = row.GetCellByColumnName("M"); // Ali Order #
                    var customerName = row.GetCellByColumnName("N"); // Name
                    var customerAddress = row.GetCellByColumnName("O"); // Address
                    var customerAddress2 = row.GetCellByColumnName("P"); // Address2
                    var customerCity = row.GetCellByColumnName("Q"); // City
                    var customerZip = row.GetCellByColumnName("R"); // Zip
                    var orderNote = row.GetCellByColumnName("V"); // Order note
                    var orderState = row.GetCellByColumnName("W"); // Order state

                    var aliOrderNumberObject = (aliOrderNumber != null ? aliOrderNumber.Value : "");
                    Console.WriteLine("{0} {1} {2} {3}", orderNumber.Value, createdDate.Value, aliOrderNumberObject, customerName.Value);
                }
            }
        }

        static void ReadAliExpressOrders(string aliExpressOrdersFilesPath)
        {
            using (TextReader fileReader = File.OpenText(aliExpressOrdersFilesPath))
            {
                fileReader.ReadLine(); // skip excel separator line

                using (var csvReader = new CsvReader(fileReader))
                {
                    csvReader.Configuration.Delimiter = ",";
                    csvReader.Configuration.HasHeaderRecord = true;

                    var aliExpressOrders = csvReader.GetRecords<AliExpressOrder>();

                    foreach (var aliExpressOrder in aliExpressOrders)
                    {
                        Console.WriteLine("{0} {1} {2} {3}", aliExpressOrder.OrderId, aliExpressOrder.OrderTime, aliExpressOrder.OrderAmount, aliExpressOrder.ContactName);
                    }
                }
            }

        }

        public class AliExpressOrder
        {
            public string OrderId { get; set; }
            public string OrderTime { get; set; }
            public string StoreName { get; set; }
            public string StoreUrl { get; set; }
            public string OrderAmount { get; set; }
            public string OrderLines { get; set; }
            public string ContactName { get; set; }
            public string ContactAddress { get; set; }
            public string contactAddress2 { get; set; }
            public string ContactZipCode { get; set; }
        }
    }
}
