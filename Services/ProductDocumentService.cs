using DRAFTINGITF;
using ProductStructureTypeLib;
namespace CATIAAssistant.Services
{
    public class ProductDocumentService
    {
        private readonly ProductDocument _productDocument;

        public ProductDocumentService(ProductDocument productDocument)
        {
            productDocument = productDocument;
        }

        public void GetProductBomParameterValues(ProductDocument productDocument)
        {
            string name = _productDocument.Product.get_Name();
            string name2 = _productDocument.Product.Products.Name;
        }
    }
}
