using Catia_Macro_Test.Services;
using CATIAAssistant.Helpers;
using CATIAAssistant.Models;
using CATIAAssistant.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace CATIAAssistant
{
    public partial class Form1 : Form
    {
        private string _docType;
        private INFITF.Application _catia;
        private INFITF.Document _activeDoc;
        private ProductStructureTypeLib.ProductDocument _productDoc;
        private MECMOD.PartDocument _partDoc;
        private DRAFTINGITF.DrawingDocument _drawingDoc;
        private readonly ValidationHelper _validationHelper;
        public Form1(ValidationHelper validationHelper)
        {
            InitializeComponent();
            _validationHelper = validationHelper;
        }

        #region Form Load and UI Initialization

        private void Form1_Load(object sender, EventArgs e)
        {
            TopMost = true;
            InformationLabel.Text = "";
            // Sabit metin: "Active document:" kýsmý her zaman görünür
            ActiveDocumentPrefixLabel.Text = "Active Catia Doc:";
            ActiveDocumentLabel.Text = "";
            // Sabit metin: "Active excel:" kýsmý her zaman görünür
            ActiveExcelPrefixLabel.Text = "Active Excel Doc:";
            ActiveExcelLabel.Text = "";

        }

        #endregion
        #region Button1 Click Handlers

        // Button1: CATIA'ya baðlan ve aktif dokümaný initialize et.
        private void button1_Click(object sender, EventArgs e)
        {
            // Temizleme iþlemleri
            ActiveDocumentLabel.Text = "";
            InformationLabel.Text = "";

            // Catia baðlantýsý
            var comService = new COMService();
            try
            {
                _catia = (INFITF.Application)comService.GetActiveObject("CATIA.Application");

            }
            catch (Exception)
            {
                ActiveDocumentLabel.ForeColor = Color.Red;
                ActiveDocumentLabel.Text = "CATIA application cannot be found";
                return;
            }

            // Yeni bir CatiaDocumentHelper oluþturuyoruz.
            var docHelper = new CatiaDocumentHelper(_catia);

            // Doküman sayýsýný kontrol ediyoruz.
            if (docHelper.GetDocumentsCount() == 0)
            {
                ActiveDocumentLabel.ForeColor = Color.Red;
                ActiveDocumentLabel.Text = "No document found";
                return;
            }

            // Doküman adýný siyah renkle ekliyoruz.
            _activeDoc = docHelper.GetActiveDocument();
            ActiveDocumentLabel.ForeColor = Color.Black;
            ActiveDocumentLabel.Text = $"{_activeDoc.get_Name()}";

            // Doküman türünü alýyoruz.
            _docType = docHelper.GetDocumentType(_activeDoc);

            // Doküman türüne göre ilgili nesneyi set ediyoruz.
            if (_docType == "DrawingDocument")
            {
                _drawingDoc = (DRAFTINGITF.DrawingDocument)_activeDoc;
            }
            else if (_docType == "ProductDocument")
            {
                _productDoc = (ProductStructureTypeLib.ProductDocument)_activeDoc;
            }
            else if (_docType == "PartDocument")
            {
                _partDoc = (MECMOD.PartDocument)_activeDoc;
            }
            else
            {
                ActiveDocumentLabel.ForeColor = Color.Red;
                ActiveDocumentLabel.Text = "Document type is not supported";
                return;
            }
        }
        #endregion
        #region Button2 Click Handlers
        // Button2: DrawingDocument içindeki component'lardan metin verilerini DataGridView'e aktar.
        private void button2_Click(object sender, EventArgs e)
        {
            InformationLabel.Text = "";

            if (!_validationHelper.ValidateDrawingDocument(_docType))
            {
                InformationLabel.Text = "Type of this document is not \"Drawing\"";
                dataGridView1.Rows.Clear();
                return;
            }

            if (!_validationHelper.ValidateSheetsCount(_drawingDoc))
            {
                InformationLabel.Text = "No sheet found in this drawing";
                dataGridView1.Rows.Clear();
                return;
            }

            if (!_validationHelper.ValidateDetailSheet(_drawingDoc))
            {
                InformationLabel.Text = "Can not read component datas in detail sheet";
                dataGridView1.Rows.Clear();
                return;
            }

            if (!_validationHelper.ValidateActiveSheetViewsCount(_drawingDoc))
            {
                InformationLabel.Text = "No view found in this sheet";
                dataGridView1.Rows.Clear();
                return;
            }
            if (!_validationHelper.ValidateActiveView(_drawingDoc))
            {
                InformationLabel.Text = "No active view found in this sheet";
                dataGridView1.Rows.Clear();
                return;
            }

            if (!_validationHelper.ValidateActiveViewComponentsCount(_drawingDoc))
            {
                InformationLabel.Text = "No component found in the active view";
                dataGridView1.Rows.Clear();
                return;
            }

            var drawingService = new DrawingDocumentService(_drawingDoc);
            var dataRows = drawingService.GetDrawingComponentsTextData();

            // DataGridView sütunlarýný, en fazla veri içeren satýrýn uzunluðuna göre sabitliyoruz.
            if (dataRows.Count > 0)
            {
                int columnCount = dataRows[0].Length;
                dataGridView1.Columns.Clear();
                for (int i = 0; i < columnCount; i++)
                {
                    dataGridView1.Columns.Add($"Column{i + 1}", $"Column{i + 1}");
                }
            }
            dataGridView1.Rows.Clear();
            foreach (var row in dataRows)
            {
                dataGridView1.Rows.Add(row);
            }
            SetRowNumber(dataGridView1);

            var catiaComponents = new List<ComponentItem>();
            var parseQuantityHelper = new ParseQuantityHelper();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                // 0. hücrede ItemNo, 1. hücrede "2x/3x" gibi bir metin varsayýyoruz
                if (row.Cells[0].Value != null && row.Cells[1].Value != null)
                {
                    string itemNo = row.Cells[0].Value.ToString();
                    string quantityText = row.Cells[1].Value.ToString(); // Örneðin "2x/3x"

                    // Slash üzerinden parçalama
                    // "2x/3x" => ["2x", "3x"]
                    string[] parts = quantityText.Split('/');

                    int quantityDrawn = 0;
                    int quantityMirror = 0;

                    // parts[0] = "2x" => quantityDrawn
                    if (parts.Length > 0)
                    {
                        quantityDrawn = parseQuantityHelper.ParseQuantity(parts[0]);
                    }

                    // parts[1] = "3x" => quantityMirror
                    if (parts.Length > 1)
                    {
                        quantityMirror = parseQuantityHelper.ParseQuantity(parts[1]);
                    }

                    catiaComponents.Add(new ComponentItem
                    {
                        ItemNo = itemNo,
                        QuantityDrawn = quantityDrawn,
                        QuantityMirror = quantityMirror
                    });
                }
            }
        }
        #endregion
        #region Button3 Click Handlers
        private void button3_Click(object sender, EventArgs e)
        {
            InformationLabel.Text = "";
            ActiveExcelLabel.ForeColor = Color.Black;

            if (!_validationHelper.ValidateDrawingDocument(_docType))
            {
                InformationLabel.Text = "Type of this document is not \"Drawing\"";
                dataGridView1.Rows.Clear();
                return;
            }

            // Excel BOM dosya yolu
            string excelPath = _drawingDoc?.FullName?.Split('.')[0] + ".xlsx";

            using (var excelService = new ExcelService())
            {
                if (!excelService.OpenWorkbook(excelPath))
                {
                    ActiveExcelLabel.ForeColor = Color.Red;
                    ActiveExcelLabel.Text = "Excel document cannot be found";
                    return;
                }

                ActiveExcelLabel.Text = $"{excelService.Workbook.Name}";

                Excel.Range usedRange = excelService.GetUsedRange();
                // Örneðin: satýr 14'ten 100'e kadar kontrol edelim.
                excelService.ProcessUsedRange(usedRange, 14, 100);
            }
        }
        #endregion
        #region DataGridView Row Numbering

        private void dataGridView1_Sorted(object sender, EventArgs e)
        {
            SetRowNumber(dataGridView1);
        }

        private void SetRowNumber(DataGridView dgv)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.HeaderCell.Value = (row.Index + 1).ToString();
            }
        }

        #endregion
        #region Other UI Handlers

        private void checkBoxAlwaysOnTop_CheckedChanged(object sender, EventArgs e)
        {
            this.TopMost = checkBoxAlwaysOnTop.Checked;
        }

        #endregion
    }
}

