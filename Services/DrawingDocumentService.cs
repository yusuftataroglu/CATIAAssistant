using CATIAAssistant.Helpers;
using DRAFTINGITF;

namespace CATIAAssistant.Services
{
    public class DrawingDocumentService
    {
        private readonly DrawingDocument _drawingDocument;

        public DrawingDocumentService(DrawingDocument drawingDocument)
        {
            _drawingDocument = drawingDocument;
        }

        /// <summary>
        /// DrawingDocument içindeki 2D component'ların metin verilerini satır bazında alır.
        /// Her satır, drawingComponent'taki tüm metinleri içeren dizi olarak döner.
        /// Eksik hücreler boş string ile doldurulur.
        /// </summary>
        /// <returns>List of string arrays, her biri bir satır verisini temsil eder.</returns>
        public List<string[]> GetDrawingComponentsTextData(bool includeOtherViews)
        {
            var dataRows = new List<string[]>();
            ValidationHelper validationHelper = new ValidationHelper();

            // Eğer includeOtherViews = false ise, yalnızca aktif view üzerinde işlem yapacağız.
            if (!includeOtherViews)
            {
                // Component sayısını kontrol et
                if (!validationHelper.ValidateActiveViewComponentsCount(_drawingDocument))
                    throw new Exception("No component found in the active view");

                // Sadece aktif view
                DrawingComponents components = _drawingDocument.Sheets.ActiveSheet.Views.ActiveView.Components;
                AppendComponentsDataRows(components, dataRows);
            }
            else
            {
                // Tüm view'ler
                bool isEmpty = true;
                DrawingViews drawingViews = _drawingDocument.Sheets.ActiveSheet.Views;
                foreach (DrawingView drawingView in drawingViews)
                {
                    string drawingViewName = drawingView.get_Name();
                    if (!validationHelper.ValidateGivenViewComponentsCount(drawingView) || drawingViewName == "Main View" || drawingViewName == "Background View")
                        continue;
                    isEmpty = false;
                    DrawingComponents components = drawingView.Components;
                    AppendComponentsDataRows(components, dataRows);
                }
                if (isEmpty)
                    throw new Exception("No component found in views");
            }

            return dataRows;
        }

        /// <summary>
        /// Belirtilen DrawingComponents koleksiyonundaki component'ları parse edip dataRows listesine ekler.
        /// </summary>
        private void AppendComponentsDataRows(DrawingComponents components, List<string[]> dataRows)
        {
            int maxColumns = GetMaxModifiableCount(components);

            foreach (DrawingComponent drawingComponent in components)
            {
                var textValues = new List<string>();
                int modifiableCount = drawingComponent.GetModifiableObjectsCount();
                bool isEmpty = true;

                for (int i = 1; i <= modifiableCount; i++)
                {
                    isEmpty = false;
                    DrawingText text = (DrawingText)drawingComponent.GetModifiableObject(i);
                    textValues.Add(text.get_Text());
                }

                // Eksik sütunları boş string ile tamamla.
                while (textValues.Count < maxColumns)
                {
                    textValues.Add(string.Empty);
                }

                if (!isEmpty)
                    dataRows.Add(textValues.ToArray());
            }
        }


        /// <summary>
        /// DrawingComponents içindeki modifiable nesnelerin en yüksek sayısını bulur.
        /// </summary>
        /// <param name="components">DrawingComponents koleksiyonu.</param>
        /// <returns>Maksimum modifiable count.</returns>
        public int GetMaxModifiableCount(DrawingComponents components)
        {
            int maxCount = 0;
            foreach (DrawingComponent comp in components)
            {
                int count = comp.GetModifiableObjectsCount();
                if (count > maxCount)
                    maxCount = count;
            }
            return maxCount;
        }
    }
}
