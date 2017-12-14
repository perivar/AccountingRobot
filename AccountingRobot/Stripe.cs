using Stripe;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccountingRobot
{
    public static class Stripe
    {
        public static List<StripeBalanceTransaction> GetTransactions(string stripeApiKey)
        {
            StripeConfiguration.SetApiKey(stripeApiKey);
            var balanceService = new StripeBalanceService();            
            var allBalanceTransactions = new List<StripeBalanceTransaction>();

            var lastId = String.Empty;

            int MAX_PAGINATION = 100;
            int itemsExpected = MAX_PAGINATION;
            while (itemsExpected == MAX_PAGINATION)
            {
                IEnumerable<StripeBalanceTransaction> balanceTransactions = null;
                if (String.IsNullOrEmpty(lastId))
                {
                    balanceTransactions = balanceService.List(
                    new StripeBalanceTransactionListOptions() { Limit = MAX_PAGINATION });
                    itemsExpected = balanceTransactions.Count();
                }
                else
                {
                    balanceTransactions = balanceService.List(
                    new StripeBalanceTransactionListOptions()
                    {
                        Limit = MAX_PAGINATION,
                        StartingAfter = lastId
                    });
                    itemsExpected = balanceTransactions.Count();
                }

                allBalanceTransactions.AddRange(balanceTransactions);
                lastId = balanceTransactions.LastOrDefault().Id;
            }

            return allBalanceTransactions;
        }
    }
}
