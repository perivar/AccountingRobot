using PayPal.PayPalAPIInterfaceService.Model;
using System;
using System.Collections.Generic;
using PayPal.PayPalAPIInterfaceService;
using System.Globalization;

namespace AccountingRobot
{
    public static class Paypal
    {
        public static List<PayPalTransaction> GetPayPalTransactions(string payPalApiUsername, string payPalApiPassword, string payPalApiSignature, DateTime from, DateTime to)
        {
            TransactionSearchReq req = new TransactionSearchReq();
            req.TransactionSearchRequest = new TransactionSearchRequestType();

            string startDate = string.Format("{0:yyyy-MM-ddTHH\\:mm\\:ssZ}", from);
            string endDate = string.Format("{0:yyyy-MM-ddTHH\\:mm\\:ssZ}", to.AddDays(1));
            req.TransactionSearchRequest.StartDate = startDate;
            req.TransactionSearchRequest.EndDate = endDate;

            Dictionary<string, string> config = new Dictionary<string, string>();
            config.Add("mode", "live");
            config.Add("account1.apiUsername", payPalApiUsername);
            config.Add("account1.apiPassword", payPalApiPassword);
            config.Add("account1.apiSignature", payPalApiSignature);

            PayPalAPIInterfaceServiceService service = new PayPalAPIInterfaceServiceService(config);

            TransactionSearchResponseType transactionSearchResponseType = service.TransactionSearch(req);
            if (transactionSearchResponseType.Ack == AckCodeType.FAILURE)
            {
                foreach (var e in transactionSearchResponseType.Errors)
                {
                    Console.WriteLine(e.LongMessage);
                }
            }

            var payPalTransactions = new List<PayPalTransaction>();
            foreach (var transaction in transactionSearchResponseType.PaymentTransactions)
            {
                var payPalTransaction = new PayPalTransaction();

                payPalTransaction.TransactionID = transaction.TransactionID;

                // Converting from paypal date to date:
                // 2017-08-30T21:13:37Z
                // var date = DateTimeOffset.Parse(paypalTransaction.Timestamp).UtcDateTime;
                payPalTransaction.Timestamp = DateTimeOffset.Parse(transaction.Timestamp).UtcDateTime;

                payPalTransaction.Status = transaction.Status;
                payPalTransaction.Type = transaction.Type;

                payPalTransaction.GrossAmount = decimal.Parse(transaction.GrossAmount.value, CultureInfo.InvariantCulture);
                payPalTransaction.NetAmount = decimal.Parse(transaction.NetAmount.value, CultureInfo.InvariantCulture);
                payPalTransaction.FeeAmount = decimal.Parse(transaction.FeeAmount.value, CultureInfo.InvariantCulture);

                if (null != transaction.Payer)
                {
                    payPalTransaction.Payer = transaction.Payer;
                }
                payPalTransaction.PayerDisplayName = transaction.PayerDisplayName;

                payPalTransactions.Add(payPalTransaction);
            }
            return payPalTransactions;
        }
    }
}
