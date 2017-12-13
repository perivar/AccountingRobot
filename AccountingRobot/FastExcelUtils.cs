using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccountingRobot
{
    public static class FastExcelUtils
    {
        /// <summary>
        /// Convert from excel date int string to actual date
        /// E.g. 39938 gets converted to 05/05/2009
        /// </summary>
        /// <param name="dateIntString">int string like 39938</param>
        /// <returns>datetime object (like 05/05/2009)</returns>
        public static DateTime GetDateFromExcelDateInt(string dateIntString)
        {
            try
            {
                double d = double.Parse(dateIntString);
                DateTime conv = DateTime.FromOADate(d);
                return conv;
            }
            catch (Exception)
            {
                return DateTime.MinValue;
            }
        }

        /// <summary>
        /// Convert from excel decimal string to a decimal
        /// </summary>
        /// <param name="currencyString">currency string like 133.3</param>
        /// <returns>decimal like 133.3</returns>
        public static decimal GetDecimalFromExcelCurrencyString(string currencyString)
        {
            //return Convert.ToDecimal(currencyString, CultureInfo.GetCultureInfo("no"));
            return Convert.ToDecimal(currencyString, CultureInfo.InvariantCulture);
        }
    }
}
