using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace AccountingRobot
{
    public static class Skandiabanken
    {
        public static SkandiabankenBankStatement ReadBankStatement(string skandiabankenTransactionsFilePath)
        {
            var skandiabankenTransactions = new List<SkandiabankenTransaction>();

            var wb = new XLWorkbook(skandiabankenTransactionsFilePath);
            var ws = wb.Worksheet("Kontoutskrift");

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
                // BOKFØRINGSDATO	
                // RENTEDATO	
                // ARKIVREFERANSE	
                // TYPE	
                // TEKST	
                // UT FRA KONTO	
                // INN PÅ KONTO
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
            var incomingBalanceLabelCell = ws.Cell(lastCellLastColumn.Address.RowNumber + 2, lastCellLastColumn.Address.ColumnNumber-2);
            var outgoingBalanceLabelCell = ws.Cell(1, lastCellLastColumn.Address.ColumnNumber-2);
            var incomingBalanceLabel = incomingBalanceLabelCell.GetString();
            var outgoingBalanceLabel = outgoingBalanceLabelCell.GetString();
            var incomingBalanceDate = ExcelUtils.GetDateFromBankStatementString(incomingBalanceLabel);
            var outgoingBalanceDate = ExcelUtils.GetDateFromBankStatementString(outgoingBalanceLabel);

            var bankStatment = new SkandiabankenBankStatement
            {
                Transactions = skandiabankenTransactions,
                IncomingBalanceDate = incomingBalanceDate,
                IncomingBalanceLabel = incomingBalanceLabel,
                IncomingBalance = incomingBalance,
                OutgoingBalanceDate = outgoingBalanceDate,
                OutgoingBalanceLabel = outgoingBalanceLabel,
                OutgoingBalance = outgoingBalance
            };

            return bankStatment;
        }
    }

    public class SkandiabankenBankStatement
    {
        public List<SkandiabankenTransaction> Transactions { get; set; }
        public DateTime IncomingBalanceDate { get; set; }
        public string IncomingBalanceLabel { get; set; }
        public decimal IncomingBalance { get; set; }
        public DateTime OutgoingBalanceDate { get; set; }
        public string OutgoingBalanceLabel { get; set; }
        public decimal OutgoingBalance { get; set; }
    }
}
