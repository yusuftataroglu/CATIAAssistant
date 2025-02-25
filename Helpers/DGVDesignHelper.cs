using CATIAAssistant.Models;

namespace CATIAAssistant.Helpers
{
    public class DGVDesignHelper
    {
        public void OrganizeDGV(DataGridView dataGridView1, List<ProductParameter> productParameters, Dictionary<string, ProductParameter> dict)
        {
            // Sütun oluşturma (örnek)
            dataGridView1.Columns.Add("ItemNo", "Item No");
            dataGridView1.Columns.Add("Index", "Index");
            dataGridView1.Columns.Add("Drawn", "Drawn");
            dataGridView1.Columns.Add("Mirror", "Mirror");
            dataGridView1.Columns.Add("Name", "Name");
            dataGridView1.Columns.Add("Supplier", "Supplier");
            dataGridView1.Columns.Add("OrderNo", "OrderNo");
            dataGridView1.Columns.Add("TypeNo", "TypeNo");
            dataGridView1.Columns.Add("CustomerOrderNo", "CustomerOrderNo");
            dataGridView1.Columns.Add("SAPNo", "SAP No");
            dataGridView1.Columns.Add("Material", "Material");
            dataGridView1.Columns.Add("Dimensions", "Dimensions");
            dataGridView1.Columns.Add("Length", "Length");
            dataGridView1.Columns.Add("Unit", "Unit");
            dataGridView1.Columns.Add("Add.", "Add.");
            dataGridView1.Columns.Add("SparePart", "SparePart");
            dataGridView1.Columns.Add("Comment", "Comment");
            dataGridView1.Columns.Add("ChildPath", "ChildPath");
            dataGridView1.Columns["ChildPath"].Visible = false;

            // Satır ekleme
            foreach (var param in productParameters)
            {
                // Karşılaştırma yapmadan önce bom listesindeki gösterime uygun hale getiriyoruz.
                int drawn = 0;
                int mirror = 0;

                string[] arrForCheckingProductIsMirrored = ["MRD", "SYM"];
                bool doesContain = false;
                foreach (string item in arrForCheckingProductIsMirrored)
                {
                    if (param.ChildPath.Contains(item))
                    {
                        mirror = param.Quantity;
                        doesContain = true;
                        break;
                    }
                }
                if (!doesContain)
                {
                    drawn = param.Quantity;
                }
                string material = "";
                if (!string.IsNullOrWhiteSpace(param.MaterialName) || !string.IsNullOrWhiteSpace(param.MaterialNo))
                {
                    material = $"{param.MaterialNo}/{param.MaterialName}";
                }


                string length = param.Length;
                if (string.IsNullOrWhiteSpace(length))
                {
                    length = param.ProfileLength;
                }

                string sparePart;
                switch (param.SparePart)
                {
                    case "S":
                        sparePart = "SPARE PART";
                        break;
                    case "W":
                        sparePart = "WEAR PART";
                        break;
                    default:
                        sparePart = "";
                        break;
                }

                string[] commentArr = { param.Comment, param.Info, param.Painting, param.HeatTreatment };

                // Filtrele (boş olmayanları al)
                var nonEmptyItems = commentArr.Where(x => !string.IsNullOrWhiteSpace(x));

                // Join
                string comment = string.Join(" / ", nonEmptyItems);


                // Düzenlemeler bitti, dgv'ye ekliyoruz.
                dataGridView1.Rows.Add(
                    param.ItemNo,
                    "0",
                    $"{drawn}x",
                    $"{mirror}x",
                    param.Name,
                    param.Supplier,
                    param.OrderNo,
                    param.TypeNo,
                    param.CustomerOrderNo,
                    "",
                    material,
                    "",
                    length,
                    "",
                    "",
                    sparePart,
                    comment,
                    param.ChildPath
                );
            }
        }
    }
}
