using DRAFTINGITF;
using INFITF;
using ProductStructureTypeLib;

namespace CATIAAssistant.Helpers
{
    public class ValidationHelper
    {
        private ProductDocument _productDocument;

        public ProductDocument ProductDocument { get => _productDocument; set => _productDocument = value; }
        /// <summary>
        /// Dokümanın bir DrawingDocument olup olmadığını kontrol eder.
        /// </summary>
        /// <returns>Geçerliyse true, aksi halde false döner.</returns>
        public bool ValidateDrawingDocument(string docType)
        {
            if (docType != "DrawingDocument")
                return false;
            return true;
        }

        /// <summary>
        /// Çizimin sheet içerip içermediğini kontrol eder.
        /// </summary>
        /// <returns>Geçerliyse true, aksi halde false döner.</returns>
        public bool ValidateSheetsCount(DrawingDocument drawingDocument)
        {
            if (drawingDocument.Sheets.Count == 0)
                return false;
            return true;
        }

        /// <summary>
        /// Aktif sheet'in view içerip içermediğini kontrol eder.
        /// </summary>
        /// <returns>Geçerliyse true, aksi halde false döner.</returns>
        public bool ValidateActiveSheetViewsCount(DrawingDocument drawingDocument)
        {
            if (drawingDocument.Sheets.ActiveSheet.Views.Count <= 2)
                return false;
            return true;
        }

        /// <summary>
        /// Aktif sheet'in aktif view içerip içermediğini kontrol eder.
        /// </summary>
        /// <returns>Geçerliyse true, aksi halde false döner.</returns>
        public bool ValidateActiveView(DrawingDocument drawingDocument)
        {
            DrawingView activeView = drawingDocument.Sheets.ActiveSheet.Views.ActiveView;
            if (activeView.get_Name() == "Main View" || activeView.get_Name() == "Background View")
                return false;
            return true;
        }

        /// <summary>
        /// Aktif view'ın component içerip içermediğini kontrol eder.
        /// </summary>
        /// <returns>Geçerliyse true, aksi halde false döner.</returns>
        public bool ValidateActiveViewComponentsCount(DrawingDocument drawingDocument)
        {
            if (drawingDocument.Sheets.ActiveSheet.Views.ActiveView.Components.Count == 0)
                return false;
            return true;
        }

        /// <summary>
        /// Verilen view'ın component içerip içermediğini kontrol eder.
        /// </summary>
        /// <returns>Geçerliyse true, aksi halde false döner.</returns>
        public bool ValidateGivenViewComponentsCount(DrawingView drawingView)
        {
            if (drawingView.Components.Count == 0)
                return false;
            return true;
        }

        /// <summary>
        /// Aktif sheet'in detail sheet olup olmadığını kontrol eder.
        /// </summary>
        /// <returns>Geçerliyse true, aksi halde false döner.</returns>
        public bool ValidateDetailSheet(DrawingDocument drawingDocument)
        {
            if (drawingDocument.Sheets.ActiveSheet.IsDetail())
                return false;
            return true;
        }

        public bool ValidateProductDocument(INFITF.Application catia, DrawingDocument drawingDocument)
        {
            try
            {
                CatiaDocumentHelper catiaDocumentHelper = new(catia);
                string drawingDocName = drawingDocument.get_Name().Split('.')[0];
                _productDocument = catia.Documents.Item(drawingDocName + ".CATProduct") as ProductDocument;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

    }
}
