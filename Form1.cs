using Catia_Macro_Test.Services;
using CATIAAssistant.Helpers;
using CATIAAssistant.Models;
using CATIAAssistant.Services;
using INFITF;
using System.Security.Cryptography;
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
        private List<ProductParameter> _catiaComponents = new List<ProductParameter>();
        private COMService _comService = new COMService();

        public Form1()
        {
            InitializeComponent();
        }

        #region Form Load and UI Initialization

        private void Form1_Load(object sender, EventArgs e)
        {
            checkBoxAlwaysOnTop.Checked = false;
            TopMost = false;
            isZSBCheckBox.Checked = true;
            InformationLabel.Text = "";
            // Sabit metin: "Active document:" kısmı her zaman görünür
            ActiveDocumentPrefixLabel.Text = "Active Catia Doc:";
            ActiveDocumentLabel.Text = "";
            // Sabit metin: "Active excel:" kısmı her zaman görünür
            ActiveExcelPrefixLabel.Text = "Active Excel Doc:";
            ActiveExcelLabel.Text = "";

        }

        #endregion
        #region Button1 Click Handlers

        // Button1: CATIA'ya bağlan ve aktif dokümanı initialize et.
        private void button1_Click(object sender, EventArgs e)
        {
            // Temizleme işlemleri
            ActiveDocumentLabel.Text = "";
            InformationLabel.Text = "";

            CatiaDocResult catiaDocResult;
            try
            {
                _catia = (INFITF.Application)_comService.GetActiveObject("CATIA.Application");

            }
            catch (Exception ex)
            {
                ActiveDocumentLabel.ForeColor = Color.Red;
                ActiveDocumentLabel.Text = ex.Message;
                return;
            }
            var docHelper = new CatiaDocumentHelper(_catia);
            try
            {
                catiaDocResult = docHelper.DoInitializeDocument();

            }
            catch (Exception ex)
            {
                ActiveDocumentLabel.ForeColor = Color.Red;
                ActiveDocumentLabel.Text = ex.Message;
                return;
            }
            // Buraya geldiyse aktif doküman vardır.
            // Doküman türünü alıyoruz.
            _docType = catiaDocResult.DocType;
            _drawingDoc = catiaDocResult.DrawingDoc;
            _productDoc = catiaDocResult.ProductDoc;
            _partDoc = catiaDocResult.PartDoc;
            _activeDoc = catiaDocResult.ActiveDoc;

            ActiveDocumentLabel.Text = _activeDoc.get_Name();
            ActiveDocumentLabel.ForeColor = Color.Black;
        }

        #endregion
        #region Button2 Click Handlers
        // Button2: DrawingDocument içindeki component'lardan metin verilerini DataGridView'e aktar.
        private void button2_Click(object sender, EventArgs e)
        {
            ActiveExcelLabel.Text = "";
            ActiveDocumentLabel.Text = "";
            InformationLabel.Text = "";
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            CatiaDocResult catiaDocResult;
            try
            {
                _catia = (INFITF.Application)_comService.GetActiveObject("CATIA.Application");

            }
            catch (Exception ex)
            {
                ActiveDocumentLabel.ForeColor = Color.Red;
                ActiveDocumentLabel.Text = ex.Message;
                return;
            }
            var docHelper = new CatiaDocumentHelper(_catia);
            try
            {
                catiaDocResult = docHelper.DoInitializeDocument();
            }
            catch (Exception ex)
            {
                ActiveDocumentLabel.ForeColor = Color.Red;
                ActiveDocumentLabel.Text = ex.Message;
                return;
            }
            if (_activeDoc is null)
            {
                ActiveDocumentLabel.ForeColor = Color.Red;
                ActiveDocumentLabel.Text = "Read document first";
                return;
            }
            Document newActive = catiaDocResult.ActiveDoc;

            if (!Equals(newActive, _activeDoc))
            {
                DialogResult dialogResult = MessageBox.Show("You are now on different document than current active document. Do you want to update active document?", "Update Document", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                // Kullanıcı mesaj kutusundaki butonlardan birine tıklamadan önce catia'yı/dokümanı kapatmış olabilir. Bunun kontrolünü yapıyoruz.
                try
                {
                    _catia = (INFITF.Application)_comService.GetActiveObject("CATIA.Application");

                }
                catch (Exception ex)
                {
                    ActiveDocumentLabel.ForeColor = Color.Red;
                    ActiveDocumentLabel.Text = ex.Message;
                    return;
                }
                docHelper = new CatiaDocumentHelper(_catia);
                try
                {
                    catiaDocResult = docHelper.DoInitializeDocument();
                }
                catch (Exception ex)
                {
                    ActiveDocumentLabel.ForeColor = Color.Red;
                    ActiveDocumentLabel.Text = ex.Message;
                    return;
                }
                if (dialogResult == DialogResult.Yes)
                {
                    // Kullanıcı Yes'e basmışsa aktif doküman güncelleniyor.
                    _activeDoc = catiaDocResult.ActiveDoc;
                    _docType = catiaDocResult.DocType;
                    _drawingDoc = catiaDocResult.DrawingDoc;
                    _productDoc = catiaDocResult.ProductDoc;
                    _partDoc = catiaDocResult.PartDoc;
                }
            }

            ActiveDocumentLabel.Text = _activeDoc.get_Name();
            ActiveDocumentLabel.ForeColor = Color.Black;

            var validationHelper = new ValidationHelper();
            if (!validationHelper.ValidateDrawingDocument(_docType))
            {
                InformationLabel.Text = "Type of this document is not \"Drawing\"";
                return;
            }

            if (!validationHelper.ValidateSheetsCount(_drawingDoc))
            {
                InformationLabel.Text = "No sheet found in this drawing";
                return;
            }

            if (!validationHelper.ValidateDetailSheet(_drawingDoc))
            {
                InformationLabel.Text = "Can not read component datas in detail sheet";
                return;
            }

            if (!validationHelper.ValidateActiveSheetViewsCount(_drawingDoc))
            {
                InformationLabel.Text = "No view found in the active sheet";
                return;
            }
            if (!validationHelper.ValidateActiveView(_drawingDoc))
            {
                InformationLabel.Text = "No active view found in the active sheet";
                return;
            }

            var drawingService = new DrawingDocumentService(_drawingDoc);
            var dataRows = new List<string[]>();
            try
            {
                dataRows = drawingService.GetDrawingComponentsTextData(checkBoxIncludeOtherViews.Checked);
            }
            catch (Exception ex)
            {
                InformationLabel.Text = ex.Message;
                return;
            }

            // Eğer component'larda okunacak veri yoksa dataRows.Count = 0 oluyor ve boşuna devam etmesini önlüyoruz.
            if (dataRows.Count == 0)
            {
                InformationLabel.Text = "No readable text found in components of active view";
                return;
            }

            // Eğer component'larda okunacak veri varsa DataGridView sütunlarını, en fazla veri içeren satırın uzunluğuna göre sabitliyoruz.
            int columnCount = dataRows[0].Length;

            for (int i = 0; i < columnCount; i++)
            {
                dataGridView1.Columns.Add($"Column{i + 1}", $"Column{i + 1}");
            }
            foreach (var row in dataRows)
            {
                dataGridView1.Rows.Add(row);
            }
            SetRowNumber(dataGridView1);
        }
        #endregion
        #region Button3 Click Handlers
        private void button3_Click(object sender, EventArgs e)
        {
            ActiveExcelLabel.Text = "";
            ActiveDocumentLabel.Text = "";
            InformationLabel.Text = "";
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            CatiaDocResult catiaDocResult;
            try
            {
                _catia = (INFITF.Application)_comService.GetActiveObject("CATIA.Application");

            }
            catch (Exception ex)
            {
                ActiveDocumentLabel.ForeColor = Color.Red;
                ActiveDocumentLabel.Text = ex.Message;
                return;
            }
            var docHelper = new CatiaDocumentHelper(_catia);
            try
            {
                catiaDocResult = docHelper.DoInitializeDocument();
            }
            catch (Exception ex)
            {
                ActiveDocumentLabel.ForeColor = Color.Red;
                ActiveDocumentLabel.Text = ex.Message;
                return;
            }
            if (_activeDoc is null)
            {
                ActiveDocumentLabel.ForeColor = Color.Red;
                ActiveDocumentLabel.Text = "Read document first";
                return;
            }
            Document newActive = catiaDocResult.ActiveDoc;

            if (!Equals(newActive, _activeDoc))
            {
                DialogResult dialogResult = MessageBox.Show("You are now on different document than current active document. Do you want to update active document?", "Update Document", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                // Kullanıcı mesaj kutusundaki butonlardan birine tıklamadan önce catia'yı/dokümanı kapatmış olabilir. Bunun kontrolünü yapıyoruz.
                try
                {
                    _catia = (INFITF.Application)_comService.GetActiveObject("CATIA.Application");

                }
                catch (Exception ex)
                {
                    ActiveDocumentLabel.ForeColor = Color.Red;
                    ActiveDocumentLabel.Text = ex.Message;
                    return;
                }
                docHelper = new CatiaDocumentHelper(_catia);
                try
                {
                    catiaDocResult = docHelper.DoInitializeDocument();
                }
                catch (Exception ex)
                {
                    ActiveDocumentLabel.ForeColor = Color.Red;
                    ActiveDocumentLabel.Text = ex.Message;
                    return;
                }
                if (dialogResult == DialogResult.Yes)
                {
                    // Kullanıcı Yes'e basmışsa aktif doküman güncelleniyor.
                    _activeDoc = catiaDocResult.ActiveDoc;
                    _docType = catiaDocResult.DocType;
                    _drawingDoc = catiaDocResult.DrawingDoc;
                    _productDoc = catiaDocResult.ProductDoc;
                    _partDoc = catiaDocResult.PartDoc;
                }
            }

            var validationHelper = new ValidationHelper();
            if (_docType != "ProductDocument")
            {
                try
                {
                    _productDoc = validationHelper.GetProductDocument(_catia, _activeDoc);
                }
                catch (Exception ex)
                {
                    ActiveDocumentLabel.ForeColor = Color.Red;
                    ActiveDocumentLabel.Text = ex.Message;
                    return;
                }
            }
            ActiveDocumentLabel.ForeColor = Color.Black;
            ActiveDocumentLabel.Text = _activeDoc.get_Name();



            //Excel BOM dosya yolu
            string documentExtensionName = _activeDoc.get_Name().Split('.')[1];
            string excelPath = $"{_activeDoc?.FullName.Replace($".{documentExtensionName}", ".xlsx")}";
            using (var excelService = new ExcelService())
            {
                try
                {
                    excelService.OpenWorkbook(excelPath);
                }
                catch (Exception ex)
                {
                    ActiveExcelLabel.ForeColor = Color.Red;
                    ActiveExcelLabel.Text = ex.Message;
                    return;
                }
                ActiveExcelLabel.ForeColor = Color.Black;
                ActiveExcelLabel.Text = $"{excelService.Workbook.Name}";

                Excel.Range usedRange = excelService.GetUsedRange();
                // Örneğin: satır 14'ten 100'e kadar kontrol edelim.
                var bomItems = excelService.ProcessUsedRange(usedRange, 14, 100);

                // Product parametrelerini alıyoruz.
                ProductDocumentService productDocumentService = new ProductDocumentService();
                productDocumentService.GetParameterValuesFromProduct(_productDoc.Product, string.Empty, isZSBCheckBox.Checked);
                List<ProductParameter> productParameters = productDocumentService.productParameters;
                Dictionary<string, ProductParameter> dict = productDocumentService._dict;

                // Sütun oluşturma (örnek)
                dataGridView1.Columns.Add("ItemNo", "Item No");
                dataGridView1.Columns.Add("Quantity", "Quantity");
                dataGridView1.Columns.Add("Name", "Name");
                dataGridView1.Columns.Add("Supplier", "Supplier");
                dataGridView1.Columns.Add("OrderNo", "OrderNo");
                dataGridView1.Columns.Add("TypeNo", "TypeNo");
                dataGridView1.Columns.Add("CustomerOrderNo", "CustomerOrderNo");
                dataGridView1.Columns.Add("Material", "Material");
                dataGridView1.Columns.Add("Dimensions", "Dimensions");
                dataGridView1.Columns.Add("Length", "Length");
                dataGridView1.Columns.Add("SparePart", "SparePart");
                dataGridView1.Columns.Add("Comment", "Comment");
                dataGridView1.Columns.Add("ChildPath", "ChildPath");
                dataGridView1.Columns["ChildPath"].Visible = false;

                // Satır ekleme
                foreach (var param in productParameters)
                {
                    string sparePart;
                    // Bom listesindeki gösterime uygun hale getiriyoruz.
                    switch (param.SparePart)
                    {
                        case "S":
                            sparePart = "SPARE PART";
                            break;
                        case "W":
                            sparePart = "WEAR PART";
                            break;
                        default:
                            sparePart = "";
                            break;
                    }
                    string material = "";
                    if (!string.IsNullOrWhiteSpace(param.MaterialName) || !string.IsNullOrWhiteSpace(param.MaterialName) || !string.IsNullOrWhiteSpace(param.MaterialNo) || !string.IsNullOrWhiteSpace(param.MaterialNo))
                    {
                        material = $"{param.MaterialNo}/{param.MaterialName}";
                    }
                    string length = param.Length;
                    if (string.IsNullOrWhiteSpace(length) || !string.IsNullOrWhiteSpace(length))
                    {
                        length = param.ProfileLength;
                    }
                    string comment = "";
                    if ((!string.IsNullOrWhiteSpace(param.Comment) || !string.IsNullOrWhiteSpace(param.Comment)) && (!string.IsNullOrWhiteSpace(param.Info) || !string.IsNullOrWhiteSpace(param.Info)))
                    {
                        comment = $"{param.Comment} / {param.Info}";
                    }
                    else if ((!string.IsNullOrWhiteSpace(param.Comment) || !string.IsNullOrWhiteSpace(param.Comment)) && (string.IsNullOrWhiteSpace(param.Info) || string.IsNullOrWhiteSpace(param.Info)))
                    {
                        comment = $"{param.Comment}";
                    }
                    else if(string.IsNullOrWhiteSpace(param.Comment) || string.IsNullOrWhiteSpace(param.Comment))
                    {
                        comment = "";
                    }

                    dataGridView1.Rows.Add(
                        param.ItemNo,
                        $"{param.Quantity}x",
                        param.Name,
                        param.Supplier,
                        param.OrderNo,
                        param.TypeNo,
                        param.CustomerOrderNo,
                        material,
                        param.Dimensions,
                        length,
                        sparePart,
                        comment,
                        param.ChildPath
                    );
                }
                SetRowNumber(dataGridView1);

                // Karşılaştırma
                ComparisonHelper comparisonHelper = new();
                comparisonHelper.CompareCatiaAndBom(dict, bomItems, dataGridView1, isZSBCheckBox.Checked);
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
        #region DataGridView Sum of Selected Cell Values
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            int totalDrawn = 0;
            int totalMirror = 0;

            // Bir HashSet ile satır indekslerini saklayacağız
            HashSet<int> selectedRowIndices = new HashSet<int>();

            foreach (DataGridViewCell cell in dataGridView1.SelectedCells)
            {
                if (cell.Value is string cellValue)
                {
                    // Quantity toplama mantığı
                    var (drawn, mirror) = new ParseQuantityHelper().ParseDrawnMirror(cellValue);
                    totalDrawn += drawn;
                    totalMirror += mirror;
                }

                // Satır indeksini ekliyoruz
                selectedRowIndices.Add(cell.RowIndex);
            }

            // Seçilen satır sayısı, HashSet'in eleman sayısı
            int rowCount = selectedRowIndices.Count;

            // InformationLabel’da hem seçilen satır sayısını hem sum değerini gösteriyoruz
            InformationLabel.Text = $"Number: {rowCount}   Sum: {totalDrawn}x/{totalMirror}x";
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

