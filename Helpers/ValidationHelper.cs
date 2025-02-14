using DRAFTINGITF;

namespace CATIAAssistant.Helpers
{
    public class ValidationHelper
    {
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
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Aktif view'ın component içerip içermediğini kontrol eder.
        /// </summary>
        /// <returns>Geçerliyse true, aksi halde false döner.</returns>
        public bool ValidateActiveViewComponentsCount(DrawingDocument drawingDocument)
        {
            if (drawingDocument.Sheets.ActiveSheet.Views.ActiveView.Components.Count == 0)
            {

                return false;
            }
            return true;
        }
    }
}
