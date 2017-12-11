using System;
using System.Globalization;
using CsvHelper;
using CsvHelper.TypeConversion;
using CsvHelper.Configuration;

namespace AccountingRobot
{
    public class AliExpressOrder
    {
        public long OrderId { get; set; }
        public DateTime OrderTime { get; set; }
        public string StoreName { get; set; }
        public string StoreUrl { get; set; }
        public decimal OrderAmount { get; set; }
        public string OrderLines { get; set; }
        public string ContactName { get; set; }
        public string ContactAddress { get; set; }
        public string ContactAddress2 { get; set; }
        public string ContactZipCode { get; set; }

        public override string ToString()
        {
            return string.Format("{0} {1} {2} {3}", OrderId, OrderTime, OrderAmount.ToString("C", new CultureInfo("en-US")), ContactName);
        }
    }

    public sealed class CsvMap : ClassMap<AliExpressOrder>
    {
        public CsvMap()
        {
            Map(m => m.OrderId);
            Map(m => m.OrderTime).TypeConverterOption.Format("HH:mm MMM. dd yyyy");
            Map(m => m.StoreName);
            Map(m => m.StoreUrl);
            Map(m => m.OrderAmount).TypeConverter<CustomCurrencyConverter>();
            Map(m => m.OrderLines);
            Map(m => m.ContactName);
            Map(m => m.ContactAddress);
            Map(m => m.ContactAddress2);
            Map(m => m.ContactZipCode);
        }
    }

    public class CustomCurrencyConverter : ITypeConverter
    {
        public object ConvertFromString(string text, IReaderRow row, MemberMapData memberMapData)
        {
            // convert string like "$ 19.80" to decimal         
            var numberFormat = new NumberFormatInfo();
            numberFormat.NegativeSign = "-";
            numberFormat.CurrencyDecimalSeparator = ".";
            numberFormat.CurrencyGroupSeparator = "";
            numberFormat.CurrencySymbol = "$ ";

            return decimal.Parse(text, NumberStyles.Currency, numberFormat);
        }

        public string ConvertToString(object value, IWriterRow row, MemberMapData memberMapData)
        {
            throw new NotImplementedException();
        }
    }
}
