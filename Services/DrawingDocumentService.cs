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

        /// <summary>
        /// DrawingDocument içindeki 2D component'ların metin verilerini satır bazında alır.
        /// Her satır, drawingComponent'taki tüm metinleri içeren dizi olarak döner.
        /// Eksik hücreler boş string ile doldurulur.
        /// </summary>
        /// <returns>List of string arrays, her biri bir satır verisini temsil eder.</returns>
        public List<string[]> GetDrawingComponentsTextData()
        {
            List<string[]> dataRows = new List<string[]>();

            // Aktif view içindeki component'ları alıyoruz.
            DrawingComponents components = _drawingDocument.Sheets.ActiveSheet.Views.ActiveView.Components;
            int maxColumns = GetMaxModifiableCount(components);

            foreach (DrawingComponent drawingComponent in components)
            {
                List<string> textValues = new List<string>();
                int modifiableCount = drawingComponent.GetModifiableObjectsCount();

                for (int i = 1; i <= modifiableCount; i++)
                {
                    DrawingText text = (DrawingText)drawingComponent.GetModifiableObject(i);
                    textValues.Add(text.get_Text());
                }

                // Eksik sütunları boş string ile tamamla.
                while (textValues.Count < maxColumns)
                {
                    textValues.Add(string.Empty);
                }
                dataRows.Add(textValues.ToArray());
            }
            return dataRows;
        }
    }
}
