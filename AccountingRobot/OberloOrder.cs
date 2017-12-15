using CsvHelper.Configuration;
using System;

namespace AccountingRobot
{
    public class OberloOrder
    {
        public string OrderNumber { get; set; }
        public DateTime CreatedDate { get; set; }
        public string FinancialStatus { get; set; }
        public string FulfillmentStatus { get; set; }
        public string Supplier { get; set; }
        public string SKU { get; set; }
        public string ProductName { get; set; }
        public string Variant { get; set; }
        public int Quantity { get; set; }
        public string TrackingNumber { get; set; }
        public string AliOrderNumber { get; set; }
        public string CustomerName { get; set; }
        public string CustomerAddress { get; set; }
        public string CustomerAddress2 { get; set; }
        public string CustomerCity { get; set; }
        public string CustomerZip { get; set; }
        public string OrderNote { get; set; }
        public string OrderState { get; set; }

        public override string ToString()
        {
            return string.Format("{0} {1:yyyy-MM-dd} {2} {3} {4}", OrderNumber, CreatedDate, AliOrderNumber, SKU, CustomerName);
        }
    }

    public sealed class OberloCsvMap : ClassMap<OberloOrder>
    {
        public OberloCsvMap()
        {
            Map(m => m.OrderNumber);
            Map(m => m.CreatedDate).TypeConverterOption.Format("dd.MM.yyyy HH.mm.ss");
            Map(m => m.FinancialStatus);
            Map(m => m.FulfillmentStatus);
            Map(m => m.Supplier);
            Map(m => m.SKU);
            Map(m => m.ProductName);
            Map(m => m.Variant);
            Map(m => m.Quantity);
            Map(m => m.AliOrderNumber);
            Map(m => m.CustomerName);
            Map(m => m.CustomerAddress);
            Map(m => m.CustomerAddress2);
            Map(m => m.CustomerCity);
            Map(m => m.CustomerZip);
        }
    }


}
