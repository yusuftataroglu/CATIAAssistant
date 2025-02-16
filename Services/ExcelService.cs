using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Catia_Macro_Test.Services
{
    public class ExcelService : IDisposable
    {
        public Excel.Application ExcelApp { get; private set; }
        public Excel.Workbook Workbook { get; private set; }
        public Excel.Worksheet Worksheet { get; private set; }

        public ExcelService()
        {
            ExcelApp = new Excel.Application();
        }

        /// <summary>
        /// Belirtilen yoldaki Excel dosyasını açar.
        /// </summary>
        public bool OpenWorkbook(string path, bool visible = true)
        {
            try
            {
                ExcelApp.Visible = visible;
                Workbook = ExcelApp.Workbooks.Open(path);
                Worksheet = Workbook.ActiveSheet as Excel.Worksheet;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Kullanılan aralığı (used range) döndürür.
        /// </summary>
        public Excel.Range GetUsedRange()
        {
            return Worksheet.UsedRange;
        }

        public void Quit()
        {
            if (Workbook != null)
            {
                Marshal.ReleaseComObject(Workbook);
            }
            if (ExcelApp != null)
            {
                ExcelApp.Quit();
                Marshal.ReleaseComObject(ExcelApp);
            }
        }

        public void Dispose()
        {
            Quit();
        }

        public void ProcessUsedRange(Excel.Range usedRange, int startRow, int endRow)
        {
            int rowCount = usedRange.Rows.Count;
            int colCount = usedRange.Columns.Count;

            for (int row = startRow; row <= endRow && row <= rowCount; row++)
            {
                bool isRowEmpty = true;
                for (int col = 1; col <= colCount; col++)
                {
                    Excel.Range cell = usedRange.Cells[row, col] as Excel.Range;
                    if (cell != null && cell.Value2 != null && !string.IsNullOrWhiteSpace(cell.Value2.ToString()))
                    {
                        isRowEmpty = false;
                        break;
                    }
                }

                if (!isRowEmpty)
                {
                    // Bu satır dolu, veriyi işleyin.
                    // Örneğin, veriyi koleksiyona ekleyin veya doğrudan DataGridView'e aktarın.
                }
            }
        }

    }
}
