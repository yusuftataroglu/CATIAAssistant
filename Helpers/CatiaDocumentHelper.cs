using Microsoft.VisualBasic;

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
            INFITF.Document document = _catia.Documents.Item(documentName);
            return document;
        }

        /// <summary>
        /// Verilen dokümanın türünü (Drawing, Part, Product) döndürür.
        /// </summary>
        public string GetDocumentType(INFITF.Document doc)
        {
            return Information.TypeName(doc);
        }

        /// <summary>
        /// CATIA'da açık olan dokümanların sayısını döndürür.
        /// </summary>
        public int GetDocumentsCount()
        {
            return _catia.Documents.Count;
        }
    }
}
