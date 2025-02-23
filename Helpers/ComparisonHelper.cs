using CATIAAssistant.Models;

namespace CATIAAssistant.Helpers
{
    public class ComparisonHelper
    {
        public void CompareCatiaAndBom(Dictionary<string, ProductParameter> productDict, List<BomItem> bomItems, DataGridView dataGridView1, bool isZSB)
        {
            Dictionary<string, BomItem> bomDict = bomItems.ToDictionary(x => x.ItemNo);

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                string itemNo = "";
                do
                {
                    itemNo = row.Cells["ItemNo"].Value?.ToString().TrimStart('0');
                }
                while (itemNo.StartsWith('0'));

                if (!bomDict.TryGetValue(itemNo, out var bomItem))
                {
                    // BOM’da yok => renklendirme
                    //foreach (DataGridViewCell cell in row.Cells)
                    //{
                    row.DefaultCellStyle.BackColor = Color.Yellow;
                    //cell.Style.BackColor = Color.Yellow;
                    //}
                }
                else
                {
                    if (row.Cells["OrderNo"].Value?.ToString() != bomItem.OrderNo)
                    {
                        // BOM’da yok => renklendirme
                        //foreach (DataGridViewCell cell in row.Cells)
                        //{
                        row.DefaultCellStyle.BackColor = Color.Orange;
                        //cell.Style.BackColor = Color.Yellow;
                        //}
                        continue;
                    }
                    // BOM’da var => Parametre karşılaştırmasını yapın
                    // Compare "Quantity" vs "QuantityDrawn"
                    string productQuantity = row.Cells["Quantity"].Value?.ToString().TrimEnd('x'); // "2x" gibi
                    string bomQuantityDrawn = bomItem.QuantityDrawn; // "2"
                    string bomQuantityMirror = bomItem.QuantityMirror; // "2"
                    string productName = row.Cells["Name"].Value?.ToString();
                    string bomName = bomItem.Description;
                    string productSupplier = row.Cells["Supplier"].Value?.ToString();
                    string bomSupplier = bomItem.Manufacturer;
                    string productOrderNo = row.Cells["OrderNo"].Value?.ToString();
                    string bomOrderNo = bomItem.OrderNo;
                    string productTypeNo = row.Cells["TypeNo"].Value?.ToString();
                    string bomTypeNo = bomItem.TypeNo;
                    string productCustomerOrderNo = row.Cells["CustomerOrderNo"].Value?.ToString();
                    string bomCustomerOrderNo = bomItem.CustomerOrderNo;
                    string productMaterial = row.Cells["Material"].Value?.ToString();
                    string bomMaterial = bomItem.MaterialNo;
                    string productDimensions = row.Cells["Dimensions"].Value?.ToString();
                    string bomDimensions = bomItem.Dimensions;
                    string productLength = row.Cells["Length"].Value?.ToString();
                    string bomLength = bomItem.Length;
                    string productSparePart = row.Cells["SparePart"].Value?.ToString();
                    string bomSparePart = bomItem.SparePart;
                    string productComment = productDict[row.Cells["ChildPath"].Value?.ToString()].Comment;
                    string bomComment = bomItem.Remark;
                    string productInfo = productDict[row.Cells["ChildPath"].Value?.ToString()].Info;



                    if (productName.Contains("SYM") || productName.Contains("MRD"))
                    {
                        if (productQuantity != bomQuantityMirror)
                        {
                            row.Cells["Quantity"].Style.BackColor = Color.Red;
                        }
                    }
                    else
                    {
                        if (productQuantity != bomQuantityDrawn)
                        {
                            row.Cells["Quantity"].Style.BackColor = Color.Red;
                        }
                    }

                    if (!productName.Equals(bomName, StringComparison.OrdinalIgnoreCase))
                    {
                        row.Cells["Name"].Style.BackColor = Color.Red;
                    }
                    if (!productSupplier.Equals(bomSupplier, StringComparison.OrdinalIgnoreCase))
                    {
                        row.Cells["Supplier"].Style.BackColor = Color.Red;
                    }
                    if (!productOrderNo.Equals(bomOrderNo, StringComparison.OrdinalIgnoreCase))
                    {
                        row.Cells["OrderNo"].Style.BackColor = Color.Red;
                    }
                    if (!productTypeNo.Equals(bomTypeNo, StringComparison.OrdinalIgnoreCase))
                    {
                        row.Cells["TypeNo"].Style.BackColor = Color.Red;
                    }
                    if (!productCustomerOrderNo.Equals(bomCustomerOrderNo, StringComparison.OrdinalIgnoreCase))
                    {
                        row.Cells["CustomerOrderNo"].Style.BackColor = Color.Red;
                    }
                    if (!productMaterial.Equals(bomMaterial, StringComparison.OrdinalIgnoreCase))
                    {
                        row.Cells["Material"].Style.BackColor = Color.Red;
                    }
                    if (!productDimensions.Equals(bomDimensions, StringComparison.OrdinalIgnoreCase))
                    {
                        row.Cells["Dimensions"].Style.BackColor = Color.Red;
                    }
                    if (!productLength.Equals(bomLength, StringComparison.OrdinalIgnoreCase))
                    {
                        row.Cells["Length"].Style.BackColor = Color.Red;
                    }
                    if (!productSparePart.Equals(bomSparePart, StringComparison.OrdinalIgnoreCase))
                    {
                        row.Cells["SparePart"].Style.BackColor = Color.Red;
                    }
                    if (!bomComment.Contains(productComment) && !bomComment.Contains(productInfo))
                    {
                        row.Cells["Comment"].Style.BackColor = Color.Red;
                    }

                }
            }

        }
    }
}

