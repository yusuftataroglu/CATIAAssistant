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
            dataGridView1 = new DataGridView();
            statusStrip1.SuspendLayout();
            statusStrip2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)splitContainer1).BeginInit();
            splitContainer1.Panel1.SuspendLayout();
            splitContainer1.Panel2.SuspendLayout();
            splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            SuspendLayout();
            // 
            // checkBoxAlwaysOnTop
            // 
            checkBoxAlwaysOnTop.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            checkBoxAlwaysOnTop.AutoSize = true;
            checkBoxAlwaysOnTop.Checked = true;
            checkBoxAlwaysOnTop.CheckState = CheckState.Checked;
            checkBoxAlwaysOnTop.Location = new Point(6, 285);
            checkBoxAlwaysOnTop.Name = "checkBoxAlwaysOnTop";
            checkBoxAlwaysOnTop.Size = new Size(129, 24);
            checkBoxAlwaysOnTop.TabIndex = 3;
            checkBoxAlwaysOnTop.Text = "Always On Top";
            checkBoxAlwaysOnTop.UseVisualStyleBackColor = true;
            checkBoxAlwaysOnTop.CheckedChanged += checkBoxAlwaysOnTop_CheckedChanged;
            // 
            // button2
            // 
            button2.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            button2.Location = new Point(6, 61);
            button2.Name = "button2";
            button2.Size = new Size(151, 49);
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
            button1.Size = new Size(151, 49);
            button1.TabIndex = 1;
            button1.Text = "Read Document";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // statusStrip1
            // 
            statusStrip1.ImageScalingSize = new Size(20, 20);
            statusStrip1.Items.AddRange(new ToolStripItem[] { InformationLabel });
            statusStrip1.Location = new Point(0, 338);
            statusStrip1.Name = "statusStrip1";
            statusStrip1.Size = new Size(699, 26);
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
            statusStrip2.ImageScalingSize = new Size(20, 20);
            statusStrip2.Items.AddRange(new ToolStripItem[] { ActiveDocumentPrefixLabel, ActiveDocumentLabel });
            statusStrip2.Location = new Point(0, 312);
            statusStrip2.Name = "statusStrip2";
            statusStrip2.Size = new Size(699, 26);
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
            splitContainer1.Location = new Point(0, 0);
            splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            splitContainer1.Panel1.Controls.Add(checkBoxAlwaysOnTop);
            splitContainer1.Panel1.Controls.Add(button1);
            splitContainer1.Panel1.Controls.Add(button2);
            // 
            // splitContainer1.Panel2
            // 
            splitContainer1.Panel2.Controls.Add(dataGridView1);
            splitContainer1.Size = new Size(699, 312);
            splitContainer1.SplitterDistance = 160;
            splitContainer1.TabIndex = 4;
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
            dataGridView1.RowHeadersWidth = 51;
            dataGridView1.Size = new Size(529, 303);
            dataGridView1.TabIndex = 0;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(699, 364);
            Controls.Add(splitContainer1);
            Controls.Add(statusStrip2);
            Controls.Add(statusStrip1);
            Name = "Form1";
            SizeGripStyle = SizeGripStyle.Show;
            Text = "Form1";
            TopMost = true;
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
    }
}
