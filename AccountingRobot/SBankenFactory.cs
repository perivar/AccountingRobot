using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using IdentityModel.Client;
using System.Configuration;
using Newtonsoft.Json;
using System.Globalization;

namespace AccountingRobot
{
    public class SBankenFactory : CachedList<SBankenTransaction>
    {
        // get SBanken configuration parameters
        string clientId = ConfigurationManager.AppSettings["SBankenApiClientId"];
        string secret = ConfigurationManager.AppSettings["SBankenApiSecret"];
        string customerId = ConfigurationManager.AppSettings["SBankenApiCustomerId"];
        string accountNumber = ConfigurationManager.AppSettings["SBankenAccountNumber"];

        public static readonly SBankenFactory Instance = new SBankenFactory();

        private SBankenFactory()
        {
        }

        protected override string CacheFileNamePrefix { get { return "SBanken Transactions"; } }

        protected override DateTime ForcedUpdateFromDate
        {
            get
            {
                return new Date().FirstDayOfTheYear;
            }
        }

        public override List<SBankenTransaction> GetCombinedUpdatedAndExisting(FileDate lastCacheFileInfo, DateTime from, DateTime to)
        {
            // we have to combine two files:
            // the original cache file and the new transactions file
            Console.Out.WriteLine("Finding SBanken transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
            var newSBankenTransactions = GetSBankenTransactions(from, to);
            var originalSBankenTransactions = Utils.ReadCacheFile<SBankenTransaction>(lastCacheFileInfo.FilePath);

            // copy all the original PayPal transactions into a new file, except entries that are 
            // from the from date or newer
            var updatedSBankenTransactions = originalSBankenTransactions.Where(p => p.TransactionDate < from).ToList();

            // and add the new transactions to beginning of list
            updatedSBankenTransactions.InsertRange(0, newSBankenTransactions);

            return updatedSBankenTransactions;
        }

        public override List<SBankenTransaction> GetList(DateTime from, DateTime to)
        {
            Console.Out.WriteLine("Finding SBanken transactions from {0:yyyy-MM-dd} to {1:yyyy-MM-dd}", from, to);
            return GetSBankenTransactions(from, to);
        }

        private List<SBankenTransaction> GetSBankenTransactions(DateTime from, DateTime to)
        {
            return GetSBankenTransactionsAsync(from, to).GetAwaiter().GetResult();
        }

        private async Task<List<SBankenTransaction>> GetSBankenTransactionsAsync(DateTime from, DateTime to)
        {
            var sBankenTransactions = new List<SBankenTransaction>();

            /** Setup constants */
            var discoveryEndpoint = "https://api.sbanken.no/identityserver";
            var apiBaseAddress = "https://api.sbanken.no";
            var bankBasePath = "/bank";
            var customersBasePath = "/customers";

            // First: get the OpenId configuration from Sbanken.
            var discoClient = new DiscoveryClient(discoveryEndpoint);

            var x = discoClient.Policy = new DiscoveryPolicy()
            {
                ValidateIssuerName = false,
            };

            var discoResult = await discoClient.GetAsync();

            if (discoResult.Error != null)
            {
                throw new Exception(discoResult.Error);
            }

            // The application now knows how to talk to the token endpoint.

            // Second: the application authenticates against the token endpoint
            var tokenClient = new TokenClient(discoResult.TokenEndpoint, clientId, secret);

            var tokenResponse = tokenClient.RequestClientCredentialsAsync().Result;

            if (tokenResponse.IsError)
            {
                throw new Exception(tokenResponse.ErrorDescription);
            }

            // The application now has an access token.

            var httpClient = new HttpClient()
            {
                BaseAddress = new Uri(apiBaseAddress),
            };

            // Finally: Set the access token on the connecting client. 
            // It will be used with all requests against the API endpoints.
            httpClient.SetBearerToken(tokenResponse.AccessToken);

            // retrieves the customer's information.
            //var customerResponse = await httpClient.GetAsync($"{customersBasePath}/api/v1/Customers/{customerId}");
            //var customerResult = await customerResponse.Content.ReadAsStringAsync();

            // retrieves the customer's accounts.
            //var accountResponse = await httpClient.GetAsync($"{bankBasePath}/api/v1/Accounts/{customerId}");
            //var accountResult = await accountResponse.Content.ReadAsStringAsync();

            // retrieve the customer's transactions
            // RFC3339 / ISO8601 with 3 decimal places
            // yyyy-MM-ddTHH:mm:ss.fffK            
            string querySuffix = string.Format(CultureInfo.InvariantCulture, "?startDate={0:yyyy-MM-ddTHH:mm:ss.fffK}&endDate={1:yyyy-MM-ddTHH:mm:ss.fffK}", from, to);
            var transactionResponse = await httpClient.GetAsync($"{bankBasePath}/api/v1/Transactions/{customerId}/{accountNumber}{querySuffix}");
            var transactionResult = await transactionResponse.Content.ReadAsStringAsync();

            // parse json
            dynamic jsonDe = JsonConvert.DeserializeObject(transactionResult);

            foreach (var transaction in jsonDe.items)
            {
                var transactionId = transaction.transactionId;
                var amount = transaction.amount;
                var text = transaction.text;
                var transactionType = transaction.transactionType;
                var accountingDate = transaction.accountingDate;
                var interestDate = transaction.interestDate;

                var sBankenTransaction = new SBankenTransaction();
                sBankenTransaction.TransactionDate = accountingDate;
                sBankenTransaction.InterestDate = interestDate;
                sBankenTransaction.ArchiveReference = transactionId;
                sBankenTransaction.Type = transactionType;
                sBankenTransaction.Text = text;

                // set account change
                sBankenTransaction.AccountChange = amount;

                if (amount > 0)
                {
                    sBankenTransaction.AccountingType = SBankenTransaction.AccountingTypeEnum.IncomeUnknown;
                }
                else
                {
                    sBankenTransaction.AccountingType = SBankenTransaction.AccountingTypeEnum.CostUnknown;
                }

                if (transactionId != null && transactionId > 0)
                {
                    sBankenTransactions.Add(sBankenTransaction);
                }
            }

            return sBankenTransactions;
        }
    }
}
