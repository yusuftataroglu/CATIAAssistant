using CATIAAssistant.Models;
using System.Diagnostics.Eventing.Reader;

namespace CATIAAssistant.Helpers
{
    public class ComparisonHelper
    {
        public void CompareCatiaAndBom(List<BomItem> bomItems, DataGridView dataGridView1, bool isZSB)
        {
            Dictionary<string, BomItem> bomDict = bomItems.ToDictionary(x => x.ItemNo);

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (!isZSB)
                {
                    string itemNo = "";
                    do
                    {
                        itemNo = row.Cells["ItemNo"].Value?.ToString().TrimStart('0');
                    }
                    while (itemNo.StartsWith('0'));

                    if (!bomDict.TryGetValue(itemNo, out var bomItem))
                    {
                        // Bom'da yok, satırı sarıya boya.
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                    }
                    else
                    {
                        if (row.Cells["OrderNo"].Value?.ToString() != bomItem.OrderNo || row.Cells["CustomerOrderNo"].Value?.ToString() != bomItem.CustomerOrderNo)
                        {
                            // Bom'da bu item no var ama order no veya customer order no farklı.
                            row.DefaultCellStyle.BackColor = Color.Orange;
                            continue;
                        }

                        // BOM’da var
                        // Compare "Quantity" vs "QuantityDrawn"
                        string productQuantityDrawn = row.Cells["Drawn"].Value?.ToString().TrimEnd('x'); // "2x" gibi
                        string productQuantityMirror = row.Cells["Mirror"].Value?.ToString().TrimEnd('x'); // "2x" gibi
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

                        string[] productCommentArr;
                        productCommentArr = row.Cells["Comment"].Value?.ToString().Split('/');
                        for (int i = 0; i < productCommentArr.Length; i++)
                            productCommentArr[i] = productCommentArr[i].Trim();
                        string bomComment = bomItem.Remark;

                        if (!productQuantityDrawn.Equals(bomQuantityDrawn, StringComparison.OrdinalIgnoreCase))
                        {
                            row.Cells["Drawn"].Style.BackColor = Color.Red;
                        }
                        if (!productQuantityMirror.Equals(bomQuantityMirror, StringComparison.OrdinalIgnoreCase))
                        {
                            row.Cells["Mirror"].Style.BackColor = Color.Red;
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
                        //if (!productLength.Equals(bomLength, StringComparison.OrdinalIgnoreCase))
                        //{
                        //    row.Cells["Length"].Style.BackColor = Color.Red;
                        //}
                        if (!productSparePart.Equals(bomSparePart, StringComparison.OrdinalIgnoreCase))
                        {
                            row.Cells["SparePart"].Style.BackColor = Color.Red;
                        }
                        bool doesContainComment = true;
                        foreach (var item in productCommentArr)
                        {
                            if (!bomComment.Contains(item))
                                doesContainComment = false;
                        }
                        if (!doesContainComment)
                        {
                            row.Cells["Comment"].Style.BackColor = Color.Red;
                        }

                    }
                }
                else
                {
                    Dictionary<string, BomItem> bomDict = bomDict = bomItems.ToDictionary(x => x.CustomerOrderNo);

                    string customerOrderNo = row.Cells["CustomerOrderNo"].Value?.ToString();

                    if (!bomDict.TryGetValue(customerOrderNo, out var bomItem))
                    {
                        // Bom'da yok, satırı sarıya boya.
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                    }
                    else
                    {
                        // BOM’da var
                        // Compare "Quantity" vs "QuantityDrawn"
                        string productQuantityDrawn = row.Cells["Drawn"].Value?.ToString().TrimEnd('x'); // "2x" gibi
                        string productQuantityMirror = row.Cells["Mirror"].Value?.ToString().TrimEnd('x'); // "2x" gibi
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

                        //string productCustomerOrderNo = row.Cells["CustomerOrderNo"].Value?.ToString();
                        //string bomCustomerOrderNo = bomItem.CustomerOrderNo;

                        string productMaterial = row.Cells["Material"].Value?.ToString();
                        string bomMaterial = bomItem.MaterialNo;

                        string productDimensions = row.Cells["Dimensions"].Value?.ToString();
                        string bomDimensions = bomItem.Dimensions;

                        string productLength = row.Cells["Length"].Value?.ToString();
                        string bomLength = bomItem.Length;

                        string productSparePart = row.Cells["SparePart"].Value?.ToString();
                        string bomSparePart = bomItem.SparePart;

                        string[] productCommentArr;
                        productCommentArr = row.Cells["Comment"].Value?.ToString().Split('/');
                        for (int i = 0; i < productCommentArr.Length; i++)
                            productCommentArr[i] = productCommentArr[i].Trim();
                        string bomComment = bomItem.Remark;

                        if (!productQuantityDrawn.Equals(bomQuantityDrawn, StringComparison.OrdinalIgnoreCase))
                        {
                            row.Cells["Drawn"].Style.BackColor = Color.Red;
                        }
                        if (!productQuantityMirror.Equals(bomQuantityMirror, StringComparison.OrdinalIgnoreCase))
                        {
                            row.Cells["Mirror"].Style.BackColor = Color.Red;
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
                        //if (!productCustomerOrderNo.Equals(bomCustomerOrderNo, StringComparison.OrdinalIgnoreCase))
                        //{
                        //    row.Cells["CustomerOrderNo"].Style.BackColor = Color.Red;
                        //}
                        if (!productMaterial.Equals(bomMaterial, StringComparison.OrdinalIgnoreCase))
                        {
                            row.Cells["Material"].Style.BackColor = Color.Red;
                        }
                        //if (!productLength.Equals(bomLength, StringComparison.OrdinalIgnoreCase))
                        //{
                        //    row.Cells["Length"].Style.BackColor = Color.Red;
                        //}
                        if (!productSparePart.Equals(bomSparePart, StringComparison.OrdinalIgnoreCase))
                        {
                            row.Cells["SparePart"].Style.BackColor = Color.Red;
                        }
                        bool doesContainComment = true;
                        foreach (var item in productCommentArr)
                        {
                            if (!bomComment.Contains(item))
                                doesContainComment = false;
                        }
                        if (!doesContainComment)
                        {
                            row.Cells["Comment"].Style.BackColor = Color.Red;
                        }

                    }
                }
            }

        }
    }
}

