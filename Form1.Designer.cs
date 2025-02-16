namespace CATIAAssistant
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            checkBoxAlwaysOnTop = new CheckBox();
            button2 = new Button();
            button1 = new Button();
            statusStrip1 = new StatusStrip();
            InformationLabel = new ToolStripStatusLabel();
            statusStrip2 = new StatusStrip();
            ActiveDocumentPrefixLabel = new ToolStripStatusLabel();
            ActiveDocumentLabel = new ToolStripStatusLabel();
            splitContainer1 = new SplitContainer();
            button3 = new Button();
            dataGridView1 = new DataGridView();
            statusStrip3 = new StatusStrip();
            ActiveExcelPrefixLabel = new ToolStripStatusLabel();
            ActiveExcelLabel = new ToolStripStatusLabel();
            statusStrip1.SuspendLayout();
            statusStrip2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)splitContainer1).BeginInit();
            splitContainer1.Panel1.SuspendLayout();
            splitContainer1.Panel2.SuspendLayout();
            splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            statusStrip3.SuspendLayout();
            SuspendLayout();
            // 
            // checkBoxAlwaysOnTop
            // 
            checkBoxAlwaysOnTop.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            checkBoxAlwaysOnTop.AutoSize = true;
            checkBoxAlwaysOnTop.Checked = true;
            checkBoxAlwaysOnTop.CheckState = CheckState.Checked;
            checkBoxAlwaysOnTop.Location = new Point(6, 270);
            checkBoxAlwaysOnTop.Name = "checkBoxAlwaysOnTop";
            checkBoxAlwaysOnTop.Size = new Size(129, 24);
            checkBoxAlwaysOnTop.TabIndex = 3;
            checkBoxAlwaysOnTop.Text = "Always On Top";
            checkBoxAlwaysOnTop.TextAlign = ContentAlignment.MiddleCenter;
            checkBoxAlwaysOnTop.UseVisualStyleBackColor = true;
            checkBoxAlwaysOnTop.CheckedChanged += checkBoxAlwaysOnTop_CheckedChanged;
            // 
            // button2
            // 
            button2.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            button2.Location = new Point(6, 61);
            button2.Name = "button2";
            button2.Size = new Size(155, 49);
            button2.TabIndex = 2;
            button2.Text = "Read Components";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // button1
            // 
            button1.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            button1.Location = new Point(6, 6);
            button1.Name = "button1";
            button1.Size = new Size(155, 49);
            button1.TabIndex = 1;
            button1.Text = "Read Document";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // statusStrip1
            // 
            statusStrip1.ImageScalingSize = new Size(20, 20);
            statusStrip1.Items.AddRange(new ToolStripItem[] { InformationLabel });
            statusStrip1.Location = new Point(0, 349);
            statusStrip1.Name = "statusStrip1";
            statusStrip1.Size = new Size(720, 26);
            statusStrip1.SizingGrip = false;
            statusStrip1.TabIndex = 1;
            statusStrip1.Text = "statusStrip1";
            // 
            // InformationLabel
            // 
            InformationLabel.Name = "InformationLabel";
            InformationLabel.Size = new Size(123, 20);
            InformationLabel.Text = "InformationLabel";
            // 
            // statusStrip2
            // 
            statusStrip2.Dock = DockStyle.Top;
            statusStrip2.ImageScalingSize = new Size(20, 20);
            statusStrip2.Items.AddRange(new ToolStripItem[] { ActiveDocumentPrefixLabel, ActiveDocumentLabel });
            statusStrip2.Location = new Point(0, 0);
            statusStrip2.Name = "statusStrip2";
            statusStrip2.Size = new Size(720, 26);
            statusStrip2.SizingGrip = false;
            statusStrip2.TabIndex = 2;
            statusStrip2.Text = "statusStrip2";
            // 
            // ActiveDocumentPrefixLabel
            // 
            ActiveDocumentPrefixLabel.Name = "ActiveDocumentPrefixLabel";
            ActiveDocumentPrefixLabel.Size = new Size(192, 20);
            ActiveDocumentPrefixLabel.Text = "ActiveDocumentPrefixLabel";
            // 
            // ActiveDocumentLabel
            // 
            ActiveDocumentLabel.Name = "ActiveDocumentLabel";
            ActiveDocumentLabel.Size = new Size(155, 20);
            ActiveDocumentLabel.Text = "ActiveDocumentLabel";
            // 
            // splitContainer1
            // 
            splitContainer1.Dock = DockStyle.Fill;
            splitContainer1.FixedPanel = FixedPanel.Panel1;
            splitContainer1.Location = new Point(0, 52);
            splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            splitContainer1.Panel1.Controls.Add(button3);
            splitContainer1.Panel1.Controls.Add(checkBoxAlwaysOnTop);
            splitContainer1.Panel1.Controls.Add(button1);
            splitContainer1.Panel1.Controls.Add(button2);
            // 
            // splitContainer1.Panel2
            // 
            splitContainer1.Panel2.Controls.Add(dataGridView1);
            splitContainer1.Size = new Size(720, 297);
            splitContainer1.SplitterDistance = 164;
            splitContainer1.TabIndex = 4;
            // 
            // button3
            // 
            button3.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            button3.Location = new Point(6, 116);
            button3.Name = "button3";
            button3.Size = new Size(155, 49);
            button3.TabIndex = 4;
            button3.Text = "Read BOM";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // dataGridView1
            // 
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new Point(3, 6);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.ReadOnly = true;
            dataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToFirstHeader;
            dataGridView1.Size = new Size(546, 288);
            dataGridView1.TabIndex = 0;
            dataGridView1.Sorted += dataGridView1_Sorted;
            // 
            // statusStrip3
            // 
            statusStrip3.Dock = DockStyle.Top;
            statusStrip3.ImageScalingSize = new Size(20, 20);
            statusStrip3.Items.AddRange(new ToolStripItem[] { ActiveExcelPrefixLabel, ActiveExcelLabel });
            statusStrip3.Location = new Point(0, 26);
            statusStrip3.Name = "statusStrip3";
            statusStrip3.Size = new Size(720, 26);
            statusStrip3.SizingGrip = false;
            statusStrip3.TabIndex = 5;
            statusStrip3.Text = "statusStrip3";
            // 
            // ActiveExcelPrefixLabel
            // 
            ActiveExcelPrefixLabel.Name = "ActiveExcelPrefixLabel";
            ActiveExcelPrefixLabel.Size = new Size(157, 20);
            ActiveExcelPrefixLabel.Text = "ActiveExcelPrefixLabel";
            // 
            // ActiveExcelLabel
            // 
            ActiveExcelLabel.Name = "ActiveExcelLabel";
            ActiveExcelLabel.Size = new Size(120, 20);
            ActiveExcelLabel.Text = "ActiveExcelLabel";
            ActiveExcelLabel.TextDirection = ToolStripTextDirection.Horizontal;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(720, 375);
            Controls.Add(splitContainer1);
            Controls.Add(statusStrip1);
            Controls.Add(statusStrip3);
            Controls.Add(statusStrip2);
            Name = "Form1";
            SizeGripStyle = SizeGripStyle.Show;
            Text = "Form1";
            TopMost = true;
            Load += Form1_Load;
            statusStrip1.ResumeLayout(false);
            statusStrip1.PerformLayout();
            statusStrip2.ResumeLayout(false);
            statusStrip2.PerformLayout();
            splitContainer1.Panel1.ResumeLayout(false);
            splitContainer1.Panel1.PerformLayout();
            splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)splitContainer1).EndInit();
            splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            statusStrip3.ResumeLayout(false);
            statusStrip3.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private Button button2;
        private Button button1;
        private StatusStrip statusStrip1;
        private ToolStripStatusLabel InformationLabel;
        private StatusStrip statusStrip2;
        private ToolStripStatusLabel ActiveDocumentPrefixLabel;
        private ToolStripStatusLabel ActiveDocumentLabel;
        private DataGridView dataGridView1;
        private CheckBox checkBoxAlwaysOnTop;
        private SplitContainer splitContainer1;
        private Button button3;
        private StatusStrip statusStrip3;
        private ToolStripStatusLabel ActiveExcelPrefixLabel;
        private ToolStripStatusLabel ActiveExcelLabel;
    }
}
