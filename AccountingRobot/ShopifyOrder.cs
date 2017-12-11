namespace AccountingRobot
{
    public class ShopifyOrder
    {
        public long Id { get; set; }
        public string Name { get; set; }
        public string FinancialStatus { get; set; }
        public string Gateway { get; set; }
        public decimal TotalPrice { get; set; }
        public decimal TotalTax { get; set; }
        public string CustomerName { get; set; }

        public override string ToString()
        {
            return string.Format("{0} {1} {2} {3} {4:C} {5:C} {6}", Id, Name, FinancialStatus, Gateway, TotalPrice, TotalTax, CustomerName);
        }
    }
}
