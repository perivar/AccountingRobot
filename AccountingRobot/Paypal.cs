using PayPal.PayPalAPIInterfaceService.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using PayPal.Api;
using PayPal.PayPalAPIInterfaceService;
using System.Configuration;

namespace AccountingRobot
{
    public static class Paypal
    {
        public static List<PaymentTransactionSearchResultType> GetTransactions()
        {
            // get paypal configuration parameters
            string payPalApiUsername = ConfigurationManager.AppSettings["PayPalApiUsername"];
            string payPalApiPassword = ConfigurationManager.AppSettings["PayPalApiPassword"];
            string payPalApiSignature = ConfigurationManager.AppSettings["PayPalApiSignature"];

            DateTime endDate = new DateTime(2017, 1, 1);
            DateTime startDate = new DateTime(2017, 12, 31);

            TransactionSearchReq req = new TransactionSearchReq();
            req.TransactionSearchRequest = new TransactionSearchRequestType();

            //req.TransactionSearchRequest.StartDate = startDate.ToString("yyyy-MM-ddTHH:mm:ssZ");
            //req.TransactionSearchRequest.EndDate = endDate.ToString("yyyy-MM-ddTHH:mm:ssZ");

            req.TransactionSearchRequest.StartDate = "2017-01-01T00:00:00Z";
            req.TransactionSearchRequest.EndDate = "2017-12-31T00:00:00Z";

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

            return transactionSearchResponseType.PaymentTransactions;
        }
    }
}
