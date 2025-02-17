using Catia_Macro_Test.Services;
using CATIAAssistant.Helpers;
using CATIAAssistant.Models;
using CATIAAssistant.Services;
using DRAFTINGITF;
using INFITF;
using Microsoft.VisualBasic;
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
        private List<ComponentItem> _catiaComponents = new List<ComponentItem>();

        public Form1()
        {
            InitializeComponent();
        }

        #region Form Load and UI Initialization

        private void Form1_Load(object sender, EventArgs e)
        {
            TopMost = true;
            InformationLabel.Text = "";
            // Sabit metin: "Active document:" k�sm� her zaman g�r�n�r
            ActiveDocumentPrefixLabel.Text = "Active Catia Doc:";
            ActiveDocumentLabel.Text = "";
            // Sabit metin: "Active excel:" k�sm� her zaman g�r�n�r
            ActiveExcelPrefixLabel.Text = "Active Excel Doc:";
            ActiveExcelLabel.Text = "";

        }

        #endregion
        #region Button1 Click Handlers

        // Button1: CATIA'ya ba�lan ve aktif dok�man� initialize et.
        private void button1_Click(object sender, EventArgs e)
        {
            // Temizleme i�lemleri
            ActiveDocumentLabel.Text = "";
            InformationLabel.Text = "";

            // Catia ba�lant�s�
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

            // Yeni bir CatiaDocumentHelper olu�turuyoruz.
            var docHelper = new CatiaDocumentHelper(_catia);

            // Dok�man say�s�n� kontrol ediyoruz.
            if (docHelper.GetDocumentsCount() == 0)
            {
                ActiveDocumentLabel.ForeColor = Color.Red;
                ActiveDocumentLabel.Text = "No document found";
                return;
            }

            // Dok�man ad�n� siyah renkle ekliyoruz.
            _activeDoc = docHelper.GetActiveDocument();
            ActiveDocumentLabel.ForeColor = Color.Black;
            ActiveDocumentLabel.Text = $"{_activeDoc.get_Name()}";

            // Dok�man t�r�n� al�yoruz.
            _docType = docHelper.GetDocumentType(_activeDoc);

            // Dok�man t�r�ne g�re ilgili nesneyi set ediyoruz.
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
        // Button2: DrawingDocument i�indeki component'lardan metin verilerini DataGridView'e aktar.
        private void button2_Click(object sender, EventArgs e)
        {
            InformationLabel.Text = "";
            var validationHelper = new ValidationHelper();
            if (!validationHelper.ValidateDrawingDocument(_docType))
            {
                InformationLabel.Text = "Type of this document is not \"Drawing\"";
                dataGridView1.Rows.Clear();
                return;
            }

            if (!validationHelper.ValidateSheetsCount(_drawingDoc))
            {
                InformationLabel.Text = "No sheet found in this drawing";
                dataGridView1.Rows.Clear();
                return;
            }

            if (!validationHelper.ValidateDetailSheet(_drawingDoc))
            {
                InformationLabel.Text = "Can not read component datas in detail sheet";
                dataGridView1.Rows.Clear();
                return;
            }

            if (!validationHelper.ValidateActiveSheetViewsCount(_drawingDoc))
            {
                InformationLabel.Text = "No view found in this sheet";
                dataGridView1.Rows.Clear();
                return;
            }
            if (!validationHelper.ValidateActiveView(_drawingDoc))
            {
                InformationLabel.Text = "No active view found in this sheet";
                dataGridView1.Rows.Clear();
                return;
            }

            if (!validationHelper.ValidateActiveViewComponentsCount(_drawingDoc))
            {
                InformationLabel.Text = "No component found in the active view";
                dataGridView1.Rows.Clear();
                return;
            }

            if (!validationHelper.ValidateProductDocument(_catia, _drawingDoc))
            {
                InformationLabel.Text = "No product found";
                dataGridView1.Rows.Clear();
                return;
            }
            // todo g�ncellenecek: sadece component say�lar�n� almak i�in kullan�labilir. kar��la�t�rma i�lemi i�in item no, product'�n properties k�sm�ndan al�nabilir.
            _productDoc = validationHelper.ProductDocument;
            var drawingService = new DrawingDocumentService(_drawingDoc);
            var productService = new ProductDocumentService(_productDoc);
            var dataRows = drawingService.GetDrawingComponentsTextData();
            productService.GetProductBomParameterValues();
            // DataGridView s�tunlar�n�, en fazla veri i�eren sat�r�n uzunlu�una g�re sabitliyoruz.
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

            var parseQuantityHelper = new ParseQuantityHelper();

            _catiaComponents.Clear();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                // 0. h�crede ItemNo, 1. h�crede "2x/3x" gibi bir metin varsay�yoruz
                if (row.Cells[0].Value != null && row.Cells[1].Value != null)
                {
                    string itemNo = row.Cells[0].Value.ToString();
                    string quantityText = row.Cells[1].Value.ToString(); // �rne�in "2x/3x"

                    // Slash �zerinden par�alama
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

                    _catiaComponents.Add(new ComponentItem
                    {
                        ItemNo = int.Parse(itemNo),// todo hata verebilir.
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
            var validationHelper = new ValidationHelper();
            if (!validationHelper.ValidateDrawingDocument(_docType))
            {
                InformationLabel.Text = "Type of this document is not \"Drawing\"";
                dataGridView1.Rows.Clear();
                return;
            }
            if (_catiaComponents.Count == 0)
            {
                InformationLabel.Text = "Read components first";
                return;
            }

            // Excel BOM dosya yolu
            string excelPath = _drawingDoc?.Path + $"{_drawingDoc?.get_Name().Split('.')[0]}.xlsx";
            ComparisonHelper comparisonHelper = new();
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
                // �rne�in: sat�r 14'ten 100'e kadar kontrol edelim.
                var bomItems = excelService.ProcessUsedRange(usedRange, 14, 100);
                // Kar��la�t�rma
                comparisonHelper.CompareCatiaAndBom(_catiaComponents, bomItems);
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

