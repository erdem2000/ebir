#nullable disable
namespace ExcelBirlestirici;

partial class Form1
{
    private System.ComponentModel.IContainer components = null;

    private TextBox txtFolder = null!;
    private Button btnBrowse = null!;
    private Button btnMerge = null!;
    private Label lblStatus = null!;
    private CheckBox chkSkipInvalidRows = null!;
    private CheckBox chkWorkbookProtectedMode = null!;
    private Label lblSheetCount = null!;
    private NumericUpDown numSheetCount = null!;
    private FolderBrowserDialog folderDialog = null!;

    protected override void Dispose(bool disposing)
    {
        if (disposing && components is not null)
            components.Dispose();
        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        components = new System.ComponentModel.Container();
        txtFolder = new TextBox();
        btnBrowse = new Button();
        btnMerge = new Button();
        lblStatus = new Label();
        chkSkipInvalidRows = new CheckBox();
        chkWorkbookProtectedMode = new CheckBox();
        lblSheetCount = new Label();
        numSheetCount = new NumericUpDown();
        folderDialog = new FolderBrowserDialog();
        ((System.ComponentModel.ISupportInitialize)numSheetCount).BeginInit();
        SuspendLayout();

        txtFolder.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        txtFolder.Location = new Point(12, 12);
        txtFolder.ReadOnly = true;
        txtFolder.Size = new Size(560, 27);
        txtFolder.TabIndex = 0;
        txtFolder.PlaceholderText = "Klasör seçin…";

        btnBrowse.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnBrowse.Location = new Point(578, 10);
        btnBrowse.Size = new Size(94, 31);
        btnBrowse.TabIndex = 1;
        btnBrowse.Text = "Klasör…";
        btnBrowse.UseVisualStyleBackColor = true;
        btnBrowse.Click += BtnBrowse_Click;

        btnMerge.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnMerge.Location = new Point(578, 136);
        btnMerge.Size = new Size(94, 34);
        btnMerge.TabIndex = 5;
        btnMerge.Text = "Birleştir";
        btnMerge.UseVisualStyleBackColor = true;
        btnMerge.Click += BtnMerge_Click;

        chkSkipInvalidRows.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        chkSkipInvalidRows.AutoSize = true;
        chkSkipInvalidRows.Location = new Point(12, 47);
        chkSkipInvalidRows.Size = new Size(260, 24);
        chkSkipInvalidRows.TabIndex = 2;
        chkSkipInvalidRows.Text = "Hatalı satırları atla ve devam et";
        chkSkipInvalidRows.UseVisualStyleBackColor = true;

        chkWorkbookProtectedMode.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        chkWorkbookProtectedMode.AutoSize = true;
        chkWorkbookProtectedMode.Checked = true;
        chkWorkbookProtectedMode.CheckState = CheckState.Checked;
        chkWorkbookProtectedMode.Location = new Point(12, 74);
        chkWorkbookProtectedMode.Size = new Size(240, 24);
        chkWorkbookProtectedMode.TabIndex = 3;
        chkWorkbookProtectedMode.Text = "Çalışma Kitabı Korumalı Mod";
        chkWorkbookProtectedMode.UseVisualStyleBackColor = true;

        lblSheetCount.AutoSize = true;
        lblSheetCount.Location = new Point(12, 105);
        lblSheetCount.Text = "Birleştirilecek sayfa sayısı (ilk N):";

        numSheetCount.Location = new Point(280, 101);
        numSheetCount.Minimum = 1;
        numSheetCount.Maximum = 100;
        numSheetCount.Size = new Size(72, 27);
        numSheetCount.TabIndex = 4;
        numSheetCount.Value = 1;

        lblStatus.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        lblStatus.Location = new Point(12, 140);
        lblStatus.Size = new Size(560, 60);
        lblStatus.TabIndex = 6;
        lblStatus.Text = "Hazır.";

        AutoScaleDimensions = new SizeF(8F, 20F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(684, 209);
        Controls.Add(numSheetCount);
        Controls.Add(lblSheetCount);
        Controls.Add(chkWorkbookProtectedMode);
        Controls.Add(chkSkipInvalidRows);
        Controls.Add(lblStatus);
        Controls.Add(btnMerge);
        Controls.Add(btnBrowse);
        Controls.Add(txtFolder);
        MinimumSize = new Size(500, 248);
        Text = "THB Excel Birleştirici";
        ((System.ComponentModel.ISupportInitialize)numSheetCount).EndInit();
        ResumeLayout(false);
        PerformLayout();
    }
}
