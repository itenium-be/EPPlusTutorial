namespace EPPlusTutorial.Util
{
    public class Sell
    {
        public string Name { get; set; }
        public int Quantity { get; set; }
        public decimal Price { get; set; }
        public decimal Total => Price * Quantity;
        public decimal? Discount { get; set; }

        public Sell(string name, decimal price, int quantity, decimal? discount = null)
        {
            Name = name;
            Quantity = quantity;
            Price = price;
            Discount = discount;
        }

        public override string ToString() => $"{Name}: {Quantity} * {Price}";
    }
}