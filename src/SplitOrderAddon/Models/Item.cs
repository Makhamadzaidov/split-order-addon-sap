namespace SplitOrderAddon.Models
{
    public class Item
    {
        public string ItemCode { get; set; }
        public string ItemName { get; set; }
        public string ItmGrp { get; set; }
        public double Quantity { get; set; }
        public double DiscountPercent { get; set; }
        public double Price { get; set; }
        public string WarehouseCode { get; set; }
    }
}
