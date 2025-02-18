using DRAFTINGITF;
using ProductStructureTypeLib;
namespace CATIAAssistant.Services
{
    public class ProductDocumentService
    {
        private readonly ProductDocument _productDocument;

        public ProductDocumentService(ProductDocument productDocument)
        {
            _productDocument = productDocument;
        }

        public void GetProductParameterValues()
        {
            List<string> parameterValues = new List<string>();
            Products products = _productDocument.Product.Products;
            foreach (Product product in products)
            {
                foreach (Product item in product.Products)
                {
                parameterValues.Add(item.get_DescriptionInst());

                }
            }
        }
    }
}
