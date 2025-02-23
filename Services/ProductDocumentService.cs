using CATIAAssistant.Models;
using ProductStructureTypeLib;

namespace CATIAAssistant.Services
{
    public class ProductDocumentService
    {
        // Bu HashSet, "fullPath_itemNo" keylerini saklayarak duplicate kontrolü yapar
        public Dictionary<string, ProductParameter> _dict = new Dictionary<string, ProductParameter>();

        // Sonuç listesini buraya ekleyeceğiz
        public List<ProductParameter> productParameters { get; set; }

        public ProductDocumentService()
        {
            productParameters = new List<ProductParameter>();
        }
        #region Get product parameters
        // Ana metot
        public void GetParameterValuesFromProduct(
            Product product,
            string currentPath,  // parent path
            bool isZSB)
        {
            // currentPath: bu product'ın parent yolunu ifade eder.
            // Örneğin ilk çağrıda currentPath = "Product1" veya boş "" olabilir.

            // Kendi path'ini oluştur (örneğin parentPath + "\\" + PartNumber)
            string thisPath = string.IsNullOrWhiteSpace(currentPath)
                ? product.get_PartNumber()
                : currentPath + "\\" + product.get_PartNumber();
            bool isUselessProduct = false;

            // isZSB => sadece bir seviye
            if (isZSB)
            {
                Products children = product.Products;
                foreach (Product child in children)
                {
                    isUselessProduct = false;
                    string childPath = thisPath + "\\" + child.get_PartNumber();
                    // Tek seviye: alt product parametresini oku
                    isUselessProduct = ExtractParametersFromProduct(child, childPath);
                    if (isUselessProduct)
                        continue;
                }
                productParameters = _dict.Values.ToList();
            }
            else
            {
                // Derine in
                Products children = product.Products;

                foreach (Product child in children)
                {
                    // childPath: Bu child'ın tam konumu
                    string childPath = thisPath + "\\" + child.get_PartNumber();

                    if (child.Products.Count == 0)
                    {
                        // "Part" olduğunu varsayıyoruz (alt product yok)
                        isUselessProduct = false;
                        isUselessProduct = ExtractParametersFromProduct(child, childPath);
                        if (isUselessProduct)
                            continue;
                    }
                    else
                    {
                        // Alt montaj => recursive çağrı
                        GetParameterValuesFromProduct(child, thisPath, false);
                    }
                }
                productParameters = _dict.Values.ToList();
            }
        }

        // Ortak işlemler
        private bool ExtractParametersFromProduct(Product child, string childPath)
        {
            string name = "";
            string itemNo = "";
            string description = "";
            string supplier = "";
            string orderNo = "";
            string typeNo = "";
            string customerOrderNo = "";
            string materialName = "";
            string materialNo = "";
            string dimensions = "";
            string length = "";
            string profileLength = "";
            string sparePart = "";
            string comment = "";
            string info = "";
            string key = childPath;

            try
            {
                name = child.Parameters.Item($"{child.get_PartNumber()}\\Properties\\NAME").ValueAsString().Trim();
                if (name == "SKELETON" || name == "FIX" || name == "MOVABLE" || name == "OPENED CONDITION")
                    return true;
            }
            catch (Exception)
            {
            }
            bool isChecked = false;
            try
            {
                itemNo = child.Parameters.Item($"{child.get_PartNumber()}\\Properties\\ITEM_NO").ValueAsString().Trim();
            }
            catch (Exception)
            {
                // ItemNo yok ise Description'a bak.
                description = child.get_DescriptionInst().Trim();
                itemNo = description;
                isChecked = true;
            }
            if (!isChecked && (string.IsNullOrEmpty(itemNo) || string.IsNullOrWhiteSpace(itemNo)))
            {
                // ItemNo var ama boş ise Description'a bak.
                description = child.get_DescriptionInst().Trim();
                itemNo = description;
            }
            // ItemNo daha önce _dict'e eklenmişse boşuna diğer parametre değerlerine bakma.
            if (_dict.TryGetValue(key, out var existingParam))
            {
                existingParam.Quantity++;
                return true;
            }

            try
            {
                supplier = child.Parameters.Item($"{child.get_PartNumber()}\\Properties\\SUPPLIER").ValueAsString().Trim();
                orderNo = child.Parameters.Item($"{child.get_PartNumber()}\\Properties\\ITEM_NO_LH").ValueAsString().Trim();
                typeNo = child.Parameters.Item($"{child.get_PartNumber()}\\Properties\\TYPE_TITLE_LH").ValueAsString().Trim();
                customerOrderNo = child.Parameters.Item($"{child.get_PartNumber()}\\Properties\\DRAWING_NO").ValueAsString().Trim();
                materialName = child.Parameters.Item($"{child.get_PartNumber()}\\Properties\\MATERIAL_NAME").ValueAsString().Trim();
                materialNo = child.Parameters.Item($"{child.get_PartNumber()}\\Properties\\MATERIAL_NO").ValueAsString().Trim();
                dimensions = child.Parameters.Item($"{child.get_PartNumber()}\\Properties\\STOCK_DIM").ValueAsString().Trim();
                length = child.Parameters.Item($"{child.get_PartNumber()}\\Properties\\LENGTH").ValueAsString().Trim();
                profileLength = child.Parameters.Item($"{child.get_PartNumber()}\\Properties\\PROFILE_LENGTH").ValueAsString().Trim();
                sparePart = child.Parameters.Item($"{child.get_PartNumber()}\\Properties\\SPARE_WEAR_PART").ValueAsString().Trim();
                comment = child.Parameters.Item($"{child.get_PartNumber()}\\Properties\\COMMENT").ValueAsString().Trim();
                info = child.Parameters.Item($"{child.get_PartNumber()}\\Properties\\ADD_INFO").ValueAsString().Trim();
            }
            catch (Exception)
            {

            }
            // Yeni
            _dict[key] = new ProductParameter
            {
                ItemNo = itemNo,
                Name = name,
                Supplier = supplier,
                OrderNo = orderNo,
                TypeNo = typeNo,
                CustomerOrderNo = customerOrderNo,
                MaterialName = materialName,
                MaterialNo = materialNo,
                Dimensions = dimensions,
                Length = length,
                ProfileLength = profileLength,
                SparePart = sparePart,
                Comment = comment,
                Info = info,
                ChildPath = childPath
            };
            return false;
        }
        #endregion
    }
}
