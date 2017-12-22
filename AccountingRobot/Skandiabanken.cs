using ClosedXML.Excel;
using FastExcel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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

                    if (accountChange > 0)
                    {
                        skandiabankenTransaction.AccountingType = SkandiabankenTransaction.AccountingTypeEnum.IncomeUnknown;
                    }
                    else
                    {
                        skandiabankenTransaction.AccountingType = SkandiabankenTransaction.AccountingTypeEnum.CostUnknown;
                    }

                    skandiabankenTransactions.Add(skandiabankenTransaction);
                }
            }

            return skandiabankenTransactions;
        }

        public static List<SkandiabankenTransaction> ReadTransactions2(string skandiabankenTransactionsFilePath)
        {
            var skandiabankenTransactions = new List<SkandiabankenTransaction>();

            var wb = new XLWorkbook(skandiabankenTransactionsFilePath);
            var ws = wb.Worksheet("Kontoutskrift");

            /*
            int rowNumber = 1;
            while (ws.Cell(++rowNumber, 1).IsEmpty()) { }
            var val = ws.Cell(rowNumber, 1).Value;
            */

            var startColumn = ws.Column(1);
            var firstCellFirstColumn = startColumn.FirstCellUsed();
            var lastCellFirstColumn = startColumn.LastCellUsed();
            var lastCellLastColumn = lastCellFirstColumn.WorksheetRow().AsRange().LastColumnUsed().LastCellUsed();

            // Get a range with the transaction data
            var transactionRange = ws.Range(firstCellFirstColumn, lastCellLastColumn).RangeUsed();

            // Treat the range as a table
            var transactionTable = transactionRange.AsTable();

            // Get the transactions
            foreach (var row in transactionTable.DataRange.Rows())
            {
                var skandiabankenTransaction = new SkandiabankenTransaction();
                skandiabankenTransaction.TransactionDate = row.Field(0).GetDateTime();
                skandiabankenTransaction.InterestDate = row.Field(1).GetDateTime();
                skandiabankenTransaction.ArchiveReference = row.Field(2).GetValue<long>();
                skandiabankenTransaction.Type = row.Field(3).GetString();
                skandiabankenTransaction.Text = row.Field(4).GetString();
                skandiabankenTransaction.OutAccount = row.Field(5).GetValue<decimal>();
                skandiabankenTransaction.InAccount = row.Field(6).GetValue<decimal>();

                // set account change
                decimal accountChange = skandiabankenTransaction.InAccount - skandiabankenTransaction.OutAccount; ;

                skandiabankenTransaction.AccountChange = accountChange;

                if (accountChange > 0)
                {
                    skandiabankenTransaction.AccountingType = SkandiabankenTransaction.AccountingTypeEnum.IncomeUnknown;
                }
                else
                {
                    skandiabankenTransaction.AccountingType = SkandiabankenTransaction.AccountingTypeEnum.CostUnknown;
                }

                skandiabankenTransactions.Add(skandiabankenTransaction);
            }

            // find the incoming and outgoing balance
            var incomingBalanceCell = ws.Cell(lastCellLastColumn.Address.RowNumber + 2, lastCellLastColumn.Address.ColumnNumber);
            var outgoingBalanceCell = ws.Cell(1, lastCellLastColumn.Address.ColumnNumber);

            decimal incomingBalance = incomingBalanceCell.GetValue<decimal>();
            decimal outgoingBalance = outgoingBalanceCell.GetValue<decimal>();

            /*
            // define table header
            var tableHeader = ws.Range(firstCellFirstColumn, ws.Cell(firstCellFirstColumn.Address.RowNumber, firstCellLastColumnNumber));

            // Move to the next row (it now has the titles)
            var transactionRow = tableHeader.FirstRowUsed();
            transactionRow = transactionRow.RowBelow();

            // Get all transactions
            while (!transactionRow.Cell(1).IsEmpty())
            {
                var transactionDate = transactionRow.Cell(1).GetDateTime();
                transactionRow = transactionRow.RowBelow();
            }
            */

            return skandiabankenTransactions;
        }
    }
}
