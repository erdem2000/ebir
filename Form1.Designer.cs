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
        folderDialog = new FolderBrowserDialog();
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
        btnMerge.Location = new Point(578, 78);
        btnMerge.Size = new Size(94, 34);
        btnMerge.TabIndex = 3;
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

        lblStatus.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        lblStatus.Location = new Point(12, 82);
        lblStatus.Size = new Size(560, 60);
        lblStatus.TabIndex = 4;
        lblStatus.Text = "Hazır.";

        AutoScaleDimensions = new SizeF(8F, 20F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(684, 151);
        Controls.Add(chkSkipInvalidRows);
        Controls.Add(lblStatus);
        Controls.Add(btnMerge);
        Controls.Add(btnBrowse);
        Controls.Add(txtFolder);
        MinimumSize = new Size(500, 190);
        Text = "THB Excel Birleştirici";
        ResumeLayout(false);
        PerformLayout();
    }
}
