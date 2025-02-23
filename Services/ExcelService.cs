using CATIAAssistant.Models;
using Microsoft.Office.Interop.Excel;
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
        public void OpenWorkbook(string path, bool visible = false)
        {
            try
            {
                ExcelApp.Visible = visible;
                ExcelApp.DisplayAlerts = false;
                Workbook = ExcelApp.Workbooks.Open(path, ReadOnly: true);
                Workbook.Activate();
                Worksheet = Workbook.ActiveSheet as Excel.Worksheet;
            }
            catch (Exception)
            {
                throw new Exception("Excel document cannot be found");
            }
        }

        /// <summary>
        /// Kullanılan aralığı (used range) döndürür.
        /// </summary>
        public Excel.Range GetUsedRange()
        {
            return Worksheet.UsedRange;
        }
        public List<BomItem> ProcessUsedRange(Excel.Range usedRange, int startRow, int endRow)
        {
            //int rowCount = usedRange.Rows.Count; Sabit satır sayısı kullanıyoruz.
            int colCount = usedRange.Columns.Count;
            var bomItems = new List<BomItem>();
            for (int row = startRow; row <= endRow; row++)
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
                    string itemNo = (usedRange.Cells[row, 1] as Excel.Range)?.Value2?.ToString().Trim() ?? "";
                    string quantityDrawn = (usedRange.Cells[row, 3] as Excel.Range)?.Value2?.ToString().Trim() ?? "";
                    string quantityMirror = (usedRange.Cells[row, 4] as Excel.Range)?.Value2?.ToString().Trim() ?? "";
                    string description = (usedRange.Cells[row, 5] as Excel.Range)?.Value2?.ToString().Trim() ?? "";
                    string manufacturer = (usedRange.Cells[row, 6] as Excel.Range)?.Value2?.ToString().Trim() ?? "";
                    string orderNo = (usedRange.Cells[row, 7] as Excel.Range)?.Value2?.ToString().Trim() ?? "";
                    string typeNo = (usedRange.Cells[row, 8] as Excel.Range)?.Value2?.ToString().Trim() ?? "";
                    string customerOrderNo = (usedRange.Cells[row, 9] as Excel.Range)?.Value2?.ToString().Trim() ?? "";
                    string materialNo = (usedRange.Cells[row, 11] as Excel.Range)?.Value2?.ToString().Trim() ?? "";
                    string dimensions = (usedRange.Cells[row, 12] as Excel.Range)?.Value2?.ToString().Trim() ?? "";
                    string length = (usedRange.Cells[row, 13] as Excel.Range)?.Value2?.ToString().Trim() ?? "";
                    string sparePart = (usedRange.Cells[row, 16] as Excel.Range)?.Value2?.ToString().Trim() ?? "";
                    string remark = (usedRange.Cells[row, 17] as Excel.Range)?.Value2?.ToString().Trim() ?? "";

                    bomItems.Add(new BomItem
                    {
                        ItemNo = itemNo,
                        QuantityDrawn = quantityDrawn,
                        QuantityMirror = quantityMirror,
                        Description = description,
                        Manufacturer = manufacturer,
                        OrderNo = orderNo,
                        TypeNo = typeNo,
                        CustomerOrderNo = customerOrderNo,
                        MaterialNo = materialNo,
                        Dimensions = dimensions,
                        Length = length,
                        SparePart = sparePart,
                        Remark = remark
                    });
                }
            }
            return bomItems;
        }


        public void Quit()
        {
            // Close the workbook
            if (ExcelApp.Workbooks.Count != 0)
            {
                Workbook.Close(SaveChanges: false);
            }

            // Close the Excel application
            ExcelApp.Quit();

            // Release the COM object
            ReleaseObject(Worksheet);
            ReleaseObject(Workbook);
            ReleaseObject(ExcelApp);
        }

        // Release the COM object
        static void ReleaseObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(obj) > 0) ;
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw new Exception("Excel application failed to close");
            }
            finally
            {
                GC.Collect();
            }
        }


        public void Dispose()
        {
            Quit();
        }


    }
}
