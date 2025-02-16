using CATIAAssistant.Models;
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
        public bool OpenWorkbook(string path, bool visible = false)
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
        public void ProcessUsedRange(Excel.Range usedRange, int startRow, int endRow)
        {
            int rowCount = usedRange.Rows.Count;
            int colCount = usedRange.Columns.Count;
            var bomItems = new List<BomItem>();
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
                    // Bu satır dolu, veri çekilecek.
                    int itemNo;
                    if (!int.TryParse((usedRange.Cells[row, 1] as Excel.Range)?.Value2?.ToString(), out itemNo))
                        itemNo = 0;

                    int indexVal;
                    if (!int.TryParse((usedRange.Cells[row, 2] as Excel.Range)?.Value2?.ToString(), out indexVal))
                        indexVal = 0;

                    int drawn;
                    if (!int.TryParse((usedRange.Cells[row, 3] as Excel.Range)?.Value2?.ToString(), out drawn))
                        drawn = 0;

                    int mirror;
                    if (!int.TryParse((usedRange.Cells[row, 4] as Excel.Range)?.Value2?.ToString(), out mirror))
                        mirror = 0;

                    bomItems.Add(new BomItem
                    {
                        ItemNo = itemNo,
                        Index = indexVal,
                        Drawn = drawn,
                        Mirror = mirror
                    });
                }
            }
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


    }
}
