namespace CATIAAssistant.Models
{
    public class BomItem
    {
        public string ItemNo { get; set; }
        /*
         * Bunları product'ın parametresinden çekemiyoruz. O yüzden saklamaya gerek yok.
        public int QuantityDrawn { get; set; } 
        public int QuantityMirror { get; set; }
        */
        public string Description { get; set; }
        public string Manufacturer { get; set; }
        public string OrderNo { get; set; }
        public string TypeNo { get; set; }
        public string CustomerOrderNo { get; set; }
        public string MaterialNo { get; set; }
        public string Dimensions { get; set; }
        public string Length { get; set; }
        public string SparePart { get; set; }
        public string Remark { get; set; }


    }
}
