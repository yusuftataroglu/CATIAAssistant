using CATIAAssistant.Models;
using KnowledgewareTypeLib;
using Microsoft.Office.Interop.Excel;
using ProductStructureTypeLib;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Xml.Linq;

namespace CATIAAssistant.Services
{
    public class ProductDocumentService
    {
        // Bu HashSet, "fullPath_itemNo" keylerini saklayarak duplicate kontrolü yapar
        private HashSet<string> _seenKeys = new HashSet<string>();

        // Sonuç listesini buraya ekleyeceğiz
        public List<ProductParameter> productParameters { get; set; }

        public ProductDocumentService()
        {
            productParameters = new List<ProductParameter>();
        }

        // Ana metot
        public void GetParameterValuesFromProduct(
            Product product,
            string currentPath,  // parent path
            bool isZSB)
        {
            // currentPath: bu product'ın parent yolunu ifade eder.
            // Örneğin ilk çağrıda currentPath = "Product1" veya boş "" olabilir.

            // Kendi path'ini oluştur (örneğin parentPath + "\\" + PartNumber)
            string thisPath = string.IsNullOrWhiteSpace(dcurrentPath)
                ? product.get_PartNumber()
                : currentPath + "\\" + product.get_PartNumber();

            // isZSB => sadece bir seviye
            if (isZSB)
            {
                Products children = product.Products;

                foreach (Product child in children)
                {
                    // Tek seviye: alt product parametresini oku
                    try
                    {
                        string itemNo = child.Parameters.Item("ITEM_NO").ValueAsString().Trim();
                        string name = child.Parameters.Item("NAME").ValueAsString().Trim();
                        string supplier = child.Parameters.Item("SUPPLIER").ValueAsString().Trim();
                        string orderNo = child.Parameters.Item("ITEM_NO_LH").ValueAsString().Trim();
                        string typeNo = child.Parameters.Item("TYPE_TITLE_LH").ValueAsString().Trim();
                        string customerOrderNo = child.Parameters.Item("DRAWING_NO").ValueAsString().Trim();
                        string materialName = child.Parameters.Item("MATERIAL_NAME").ValueAsString().Trim();
                        string materialNo = child.Parameters.Item("MATERIAL_NO").ValueAsString().Trim();
                        string dimensions = child.Parameters.Item("STOCK_DIM").ValueAsString().Trim();
                        string length = child.Parameters.Item("LENGTH").ValueAsString().Trim();
                        string sparePart = child.Parameters.Item("SPARE_WEAR_PART").ValueAsString().Trim();
                        string comment = child.Parameters.Item("COMMENT").ValueAsString().Trim();
                        string info = child.Parameters.Item("ADD_INFO").ValueAsString().Trim();

                        // Key oluştur: tamPath + "_" + itemNo
                        string childPath = thisPath + "\\" + child.get_PartNumber();
                        string key = childPath + "_" + itemNo;

                        if (!_seenKeys.Contains(key))
                        {
                            _seenKeys.Add(key);
                            productParameters.Add(new ProductParameter
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
                                SparePart = sparePart,
                                Comment = comment,
                                Info = info,
                            });
                        }
                    }
                    catch (Exception)
                    {
                        // Parametre yok
                    }
                }
            }
            else
            {
                // Derine in
                Products children = product.Products;
                int childCount = children.Count;

                for (int i = 1; i <= childCount; i++)
                {
                    Product child = children.Item(i);

                    // childPath: Bu child'ın tam konumu
                    string childPath = thisPath + "\\" + child.get_PartNumber();

                    if (child.Products.Count == 0)
                    {
                        // "Part" olduğunu varsayıyoruz (alt product yok)
                        string itemNo = "";
                        string name = "";
                        string supplier = "";
                        string orderNo = "";
                        string typeNo = "";
                        string customerOrderNo = "";
                        string materialName = "";
                        string materialNo = "";
                        string dimensions = "";
                        string length = "";
                        string sparePart = "";
                        string comment = "";
                        string info = "";
                        string description = "";
                        try
                        {
                            itemNo = child.Parameters.Item("ITEM_NO").ValueAsString().Trim();
                        }
                        catch (Exception)
                        {
                            // ItemNo yok ise Description'a bak.
                            description = child.get_DescriptionInst().Trim();
                            itemNo = description;
                        }
                        if (string.IsNullOrEmpty(itemNo) || string.IsNullOrWhiteSpace(itemNo))
                        {
                            // ItemNo var ama boş ise Description'a bak.
                            description = child.get_DescriptionInst().Trim();
                            itemNo = description;
                        }
                        try
                        {
                            name = child.Parameters.Item("NAME").ValueAsString().Trim();
                            supplier = child.Parameters.Item("SUPPLIER").ValueAsString().Trim();
                            orderNo = child.Parameters.Item("ITEM_NO_LH").ValueAsString().Trim();
                            typeNo = child.Parameters.Item("TYPE_TITLE_LH").ValueAsString().Trim();
                            customerOrderNo = child.Parameters.Item("DRAWING_NO").ValueAsString().Trim();
                            materialName = child.Parameters.Item("MATERIAL_NAME").ValueAsString().Trim();
                            materialNo = child.Parameters.Item("MATERIAL_NO").ValueAsString().Trim();
                            dimensions = child.Parameters.Item("STOCK_DIM").ValueAsString().Trim();
                            length = child.Parameters.Item("LENGTH").ValueAsString().Trim();
                            sparePart = child.Parameters.Item("SPARE_WEAR_PART").ValueAsString().Trim();
                            comment = child.Parameters.Item("COMMENT").ValueAsString().Trim();
                            info = child.Parameters.Item("ADD_INFO").ValueAsString().Trim();
                        }
                        catch (Exception)
                        {
                            
                        }
                        string key = childPath + "_" + itemNo;
                        if (!_seenKeys.Contains(key))
                        {
                            _seenKeys.Add(key);
                            productParameters.Add(new ProductParameter
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
                                SparePart = sparePart,
                                Comment = comment,
                                Info = info,
                            });
                        }
                    }
                    else
                    {
                        // Alt montaj => recursive çağrı
                        GetParameterValuesFromProduct(child, thisPath, false);
                    }
                }
            }
        }
    }
}
