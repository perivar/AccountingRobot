using CsvHelper;
using FastExcel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace AccountingRobot
{
    public static class Oberlo
    {
        public static List<OberloOrder> ReadOrders(string oberloOrdersFilePath)
        {
            var oberloOrders = new List<OberloOrder>();

            // Get the input file paths
            FileInfo inputFile = new FileInfo(oberloOrdersFilePath);

            // Create a worksheet
            Worksheet worksheet = null;

            // Create an instance of Fast Excel
            using (var fastExcel = new FastExcel.FastExcel(inputFile, true))
            {
                // Read the rows using worksheet name
                string worksheetName = "Orders";
                worksheet = fastExcel.Read(worksheetName, 1);

                Console.WriteLine("Reading worksheet {0} ...", worksheetName);

                bool first = true;
                foreach (var row in worksheet.Rows)
                {
                    // skip header row
                    if (first)
                    {
                        first = false;
                        continue;
                    }

                    // read value rows
                    var orderNumber = row.GetCellByColumnName("A").Value.ToString(); // Order #
                    var createdDate = row.GetCellByColumnName("B").Value.ToString(); // Created
                    var financialStatus = row.GetCellByColumnName("C").Value.ToString(); // Financial status                    
                    var fulfillmentStatus = row.GetCellByColumnName("D").Value.ToString(); // Fulfillment status                    
                    var supplier = row.GetCellByColumnName("G").Value.ToString(); // Supplier
                    var SKU = row.GetCellByColumnName("H").Value.ToString(); // SKU
                    var productName = row.GetCellByColumnName("I").Value.ToString(); // Product
                    var variant = row.GetCellByColumnName("J").Value.ToString(); // Variant
                    var quantity = row.GetCellByColumnName("K").Value.ToString(); // Quantity
                    var trackingNumber = row.GetCellByColumnName("L"); // Tracking number
                    var aliOrderNumber = row.GetCellByColumnName("M"); // Ali Order #
                    var customerName = row.GetCellByColumnName("N").Value.ToString(); // Name
                    var customerAddress = row.GetCellByColumnName("O").Value.ToString(); // Address
                    var customerAddress2 = row.GetCellByColumnName("P"); // Address2
                    var customerCity = row.GetCellByColumnName("Q").Value.ToString(); // City
                    var customerZip = row.GetCellByColumnName("R").Value.ToString(); // Zip
                    var orderNote = row.GetCellByColumnName("V"); // Order note
                    var orderState = row.GetCellByColumnName("W").Value.ToString(); // Order state

                    var created = DateTime.ParseExact(createdDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                    var trackingNumberString = (trackingNumber != null ? trackingNumber.Value.ToString() : "");
                    var aliOrderNumberString = (aliOrderNumber != null ? aliOrderNumber.Value.ToString() : "");
                    var customerAddress2String = (customerAddress2 != null ? customerAddress2.Value.ToString() : "");
                    var orderNoteString = (orderNote != null ? orderNote.Value.ToString() : "");

                    var oberloOrder = new OberloOrder();
                    oberloOrder.OrderNumber = orderNumber;
                    oberloOrder.CreatedDate = created;
                    oberloOrder.FinancialStatus = financialStatus;
                    oberloOrder.FulfillmentStatus = fulfillmentStatus;
                    oberloOrder.Supplier = supplier;
                    oberloOrder.SKU = SKU;
                    oberloOrder.ProductName = productName;
                    oberloOrder.Variant = variant;
                    oberloOrder.Quantity = int.Parse(quantity);
                    oberloOrder.TrackingNumber = trackingNumberString;
                    oberloOrder.AliOrderNumber = aliOrderNumberString;
                    oberloOrder.CustomerName = customerName;
                    oberloOrder.CustomerAddress = customerAddress;
                    oberloOrder.CustomerAddress2 = customerAddress2String;
                    oberloOrder.CustomerCity = customerCity;
                    oberloOrder.CustomerZip = customerZip;
                    oberloOrder.OrderNote = orderNoteString;
                    oberloOrder.OrderState = orderState;

                    oberloOrders.Add(oberloOrder);
                }
            }

            return oberloOrders;
        }

        public static List<OberloOrder> ReadOrdersCSV(string oberloOrdersFilePath)
        {
            using (TextReader fileReader = File.OpenText(oberloOrdersFilePath))
            {
                using (var csvReader = new CsvReader(fileReader))
                {
                    csvReader.Configuration.Delimiter = ",";
                    csvReader.Configuration.HasHeaderRecord = true;
                    csvReader.Configuration.CultureInfo = CultureInfo.InvariantCulture;
                    csvReader.Configuration.RegisterClassMap<OberloCsvMap>();

                    return csvReader.GetRecords<OberloOrder>().ToList<OberloOrder>();
                }
            }
        }
    }
}
