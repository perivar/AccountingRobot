using FastExcel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccountingRobot
{
    public static class Skandiabanken
    {
        public static List<SkandiabankenTransaction> ReadTransactions(string skandiabankenTransactionsFilePath)
        {
            var skandiabankenTransactions = new List<SkandiabankenTransaction>();

            // Get the input file paths
            FileInfo inputFile = new FileInfo(skandiabankenTransactionsFilePath);

            // Create a worksheet
            Worksheet worksheet = null;

            // Create an instance of Fast Excel
            using (var fastExcel = new FastExcel.FastExcel(inputFile, true))
            {
                // Read the rows using worksheet name
                string worksheetName = "Kontoutskrift";
                worksheet = fastExcel.Read(worksheetName);

                Console.WriteLine("Reading worksheet {0} ...", worksheetName);

                // skip the three first rows since they only contain incoming balance and headers
                foreach (var row in worksheet.Rows.Skip(3))
                {
                    // read value rows
                    // BOKFØRINGSDATO	
                    // RENTEDATO	
                    // ARKIVREFERANSE	
                    // TYPE	
                    // TEKST	
                    // UT FRA KONTO	
                    // INN PÅ KONTO
                    var tmpValue = row.GetCellByColumnName("A").Value;

                    // if the first column (BOKFØRINGSDATO) field is empty we have reached the end 
                    // or not yet the start: The start is dealt with with the worksheet.Rows.Skip(3) command 
                    if (tmpValue.Equals(""))
                    {
                        break;
                    }
                    var transactionDateString = tmpValue.ToString();
                    var interestDateString = row.GetCellByColumnName("B").Value.ToString();
                    var archiveReferenceString = row.GetCellByColumnName("C").Value.ToString();
                    var type = row.GetCellByColumnName("D").Value.ToString();
                    var text = row.GetCellByColumnName("E").Value.ToString();
                    var outAccountString = row.GetCellByColumnName("F").Value.ToString();
                    var inAccountString = row.GetCellByColumnName("G").Value.ToString();

                    // convert to correct types
                    var transactionDate = FastExcelUtils.GetDateFromExcelDateInt(transactionDateString);
                    var interestDate = FastExcelUtils.GetDateFromExcelDateInt(interestDateString);
                    var archiveReference = long.Parse(archiveReferenceString);
                    decimal outAccount = FastExcelUtils.GetDecimalFromExcelCurrencyString(outAccountString);
                    decimal inAccount = FastExcelUtils.GetDecimalFromExcelCurrencyString(inAccountString);

                    // set account change
                    decimal accountChange = inAccount - outAccount;

                    var skandiabankenTransaction = new SkandiabankenTransaction();
                    skandiabankenTransaction.TransactionDate = transactionDate;
                    skandiabankenTransaction.InterestDate = interestDate;
                    skandiabankenTransaction.ArchiveReference = archiveReference;
                    skandiabankenTransaction.Type = type;
                    skandiabankenTransaction.Text = text;
                    skandiabankenTransaction.OutAccount = outAccount;
                    skandiabankenTransaction.InAccount = inAccount;
                    skandiabankenTransaction.AccountChange = accountChange;

                    skandiabankenTransactions.Add(skandiabankenTransaction);
                }
            }

            return skandiabankenTransactions;
        }

    }
}
