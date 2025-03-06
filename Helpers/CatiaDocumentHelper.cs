using CATIAAssistant.Models;

namespace CATIAAssistant.Helpers
{
    public class CatiaDocumentHelper
    {
        private readonly INFITF.Application _catia;

        public CatiaDocumentHelper(INFITF.Application catia)
        {
            _catia = catia;
        }

        /// <summary>
        /// CATIA'da aktif olan dokümanı döndürür.
        /// </summary>
        public INFITF.Document GetActiveDocument()
        {
            try
            {
                return _catia.ActiveDocument;
            }
            catch (Exception)
            {
                throw new Exception("No active document found");
            }
        }

        /// <summary>
        /// CATIA'da ismi verilen dokümanı döndürür.
        /// </summary>
        public INFITF.Document GetDocumentByName(string documentName)
        {
            try
            {
                INFITF.Document document = _catia.Documents.Item(documentName);
                return document;
            }
            catch (Exception)
            {
                throw new Exception("No document found with given name");
            }
        }

        /// <summary>
        /// CATIA'da açık olan dokümanların sayısını döndürür.
        /// </summary>
        public int GetDocumentsCount()
        {
            if (_catia.Documents.Count == 0)
                throw new Exception("No document found");
            return _catia.Documents.Count;
        }

        /// <summary>
        /// CATIA'da açık olan aktif dokümanın türünü döndürür.
        /// </summary>
        public CatiaDocResult GetDocAndType(INFITF.Document activeDoc)
        {
            // Doküman türünü bul
            string docType = Microsoft.VisualBasic.Information.TypeName(activeDoc);
            var result = new CatiaDocResult { DocType = docType, ActiveDoc = activeDoc };

            switch (docType)
            {
                case "DrawingDocument":
                    result.DrawingDoc = (DRAFTINGITF.DrawingDocument)activeDoc;
                    break;
                case "ProductDocument":
                    result.ProductDoc = (ProductStructureTypeLib.ProductDocument)activeDoc;
                    break;
                case "PartDocument":
                    result.PartDoc = (MECMOD.PartDocument)activeDoc;
                    break;
                default:
                    throw new Exception("Document type is not supported");
            }
            return result;
        }

        public CatiaDocResult DoInitializeDocument()
        {
            // 1) Doküman sayısını kontrol et
            GetDocumentsCount(); // hata alırsa exception fırlatır

            // 2) Aktif dokümanı al
            var activeDoc = GetActiveDocument(); // hata alırsa exception fırlatır

            // 3) Doküman tipini belirle
            var docResult = GetDocAndType(activeDoc); // hata alırsa exception fırlatır

            // docResult içinde docType, activeDoc, DrawingDoc, ProductDoc, PartDoc alanları dolu
            // docResult'ı döndür
            return docResult;
        }

    }
}
