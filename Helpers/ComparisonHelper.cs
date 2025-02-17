using CATIAAssistant.Models;

namespace CATIAAssistant.Helpers
{
    public class ComparisonHelper
    {
        public void CompareCatiaAndBom(List<ComponentItem> catiaComponents, List<BomItem> bomItems)
        {
            // BOM listesi sözlük haline getirilir (ItemNo -> BomItem)
            var bomDict = bomItems.ToDictionary(b => b.ItemNo);

            foreach (var comp in catiaComponents)
            {
                if (bomDict.TryGetValue(comp.ItemNo, out var bom))
                {
                    // Karşılaştırma
                    // 1) Quantity
                    // 2) QuantityDrawn vs bom.QuantityDrawn
                    // 3) QuantityMirror vs bom.QuantityMirror
                    if (comp.QuantityDrawn != bom.QuantityDrawn)
                    {
                        Console.WriteLine($"ItemNo={comp.ItemNo} - Drawn mismatch: CATIA={comp.QuantityDrawn}, BOM={bom.QuantityDrawn}");
                    }
                    if (comp.QuantityMirror != bom.QuantityMirror)
                    {
                        Console.WriteLine($"ItemNo={comp.ItemNo} - Mirror mismatch: CATIA={comp.QuantityMirror}, BOM={bom.QuantityMirror}");
                    }
                }
                else
                {
                    // BOM'da yok
                    Console.WriteLine($"ItemNo={comp.ItemNo} not found in BOM!");
                }
            }
        }

    }
}
