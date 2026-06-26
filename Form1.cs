using System.Text;

namespace ExcelBirlestirici;

public partial class Form1 : Form
{
    public Form1()
    {
        InitializeComponent();
    }

    private void BtnBrowse_Click(object? sender, EventArgs e)
    {
        folderDialog.InitialDirectory = string.IsNullOrWhiteSpace(txtFolder.Text)
            ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            : txtFolder.Text;

        if (folderDialog.ShowDialog(this) == DialogResult.OK)
        {
            txtFolder.Text = folderDialog.SelectedPath;
            lblStatus.Text = "Klasör seçildi. Birleştir’e basın.";
        }
    }

    private void BtnMerge_Click(object? sender, EventArgs e)
    {
        var path = txtFolder.Text.Trim();
        if (string.IsNullOrEmpty(path))
        {
            MessageBox.Show(this, "Önce bir klasör seçin.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        btnMerge.Enabled = false;
        btnBrowse.Enabled = false;
        chkSkipInvalidRows.Enabled = false;
        chkWorkbookProtectedMode.Enabled = false;
        numSheetCount.Enabled = false;

        try
        {
            var options = new MergeOptions
            {
                SkipInvalidRows = chkSkipInvalidRows.Checked,
                SheetsToMerge = (int)numSheetCount.Value,
                WorkbookProtectedMode = chkWorkbookProtectedMode.Checked,
                RequestPassword = PromptForExcelPassword
            };

            lblStatus.Text = "Sütun başlıkları kontrol ediliyor…";
            Application.DoEvents();

            var headerValidation = ExcelMergeService.ValidateSheetHeaders(path, options);
            if (!headerValidation.IsValid)
            {
                var validationMessage = string.Join(
                    Environment.NewLine + Environment.NewLine,
                    headerValidation.Errors);

                lblStatus.Text = "Sütun başlıkları uyuşmuyor. Birleştirme başlatılmadı.";
                ShowScrollablePopup(
                    "Tüm sayfalardaki sütun başlıkları aynı olmalıdır. Birleştirme başlatılmadı." +
                    Environment.NewLine + Environment.NewLine +
                    validationMessage,
                    Text,
                    MessageBoxIcon.Warning);
                return;
            }

            lblStatus.Text = "Birleştiriliyor…";
            Application.DoEvents();

            var result = ExcelMergeService.Merge(path, options);
            var sb = new StringBuilder();
            var details = new StringBuilder();

            if (result.OutputPath is not null)
            {
                sb.AppendLine($"Tamamlandı: {result.OutputPath}");
                sb.AppendLine($"İşlenen dosya: {result.FilesProcessed}, satır: {result.RowsWritten}.");
                sb.AppendLine($"PhoneNr tekrarları nedeniyle atlanan satır: {result.DuplicatesSkipped}.");
                sb.AppendLine($"Düzeltilen satır: {result.FixedRows.Count}.");
                sb.AppendLine($"Hatalı satır (atlanan): {result.SkippedRows.Count}.");
            }
            else if (result.AbortedDueToInvalidRow)
            {
                sb.AppendLine("İşlem hatalı satır nedeniyle durduruldu.");
                sb.AppendLine($"Düzeltilen satır: {result.FixedRows.Count}.");
                sb.AppendLine($"Hatalı satır (atlanan): {result.SkippedRows.Count}.");
            }
            else if (result.AbortedDueToHeaderMismatch)
            {
                sb.AppendLine("Sütun başlıkları uyuşmadığı için birleştirme başlatılmadı.");
            }

            details.AppendLine("Bulunan dosyalar:");
            if (result.DiscoveredFiles.Count == 0)
            {
                details.AppendLine("• (yok)");
            }
            else
            {
                foreach (var file in result.DiscoveredFiles)
                    details.AppendLine($"• {file}");
            }

            details.AppendLine();
            details.AppendLine("Birleştirilen dosyalar:");
            if (result.MergedFiles.Count == 0)
            {
                details.AppendLine("• (yok)");
            }
            else
            {
                foreach (var file in result.MergedFiles)
                    details.AppendLine($"• {file}");
            }

            AppendRowIssues(details, "Düzeltilen satırlar", result.FixedRows);
            AppendRowIssues(details, "Atlanan satırlar", result.SkippedRows);

            if (result.FileErrors.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Uyarı / hata:");
                foreach (var err in result.FileErrors)
                    sb.AppendLine($"• {err}");
            }

            if (result.OutputPath is null && result.FileErrors.Count == 0)
                sb.AppendLine("Çıktı oluşturulamadı.");

            lblStatus.Text = sb.ToString().Trim();

            if (result.OutputPath is not null)
            {
                var popupText = new StringBuilder();
                popupText.AppendLine($"Dosya kaydedildi:\n{result.OutputPath}");
                popupText.AppendLine();
                popupText.Append(details);
                if (result.FileErrors.Count > 0)
                {
                    popupText.AppendLine();
                    popupText.AppendLine("Uyarı / hata:");
                    foreach (var err in result.FileErrors)
                        popupText.AppendLine($"• {err}");
                }

                ShowScrollablePopup(popupText.ToString().Trim(), Text, MessageBoxIcon.Information);
            }
            else if (result.AbortedDueToInvalidRow || result.AbortedDueToHeaderMismatch || result.FileErrors.Count > 0)
            {
                var popupText = new StringBuilder();
                popupText.AppendLine(lblStatus.Text);
                popupText.AppendLine();
                popupText.Append(details);
                var icon = result.AbortedDueToInvalidRow ? MessageBoxIcon.Error : MessageBoxIcon.Warning;
                ShowScrollablePopup(popupText.ToString().Trim(), Text, icon);
            }
        }
        catch (Exception ex)
        {
            lblStatus.Text = ex.Message;
            MessageBox.Show(this, ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            btnMerge.Enabled = true;
            btnBrowse.Enabled = true;
            chkSkipInvalidRows.Enabled = true;
            chkWorkbookProtectedMode.Enabled = true;
            numSheetCount.Enabled = true;
        }
    }

    private static void AppendRowIssues(StringBuilder details, string title, IReadOnlyList<RowIssue> issues)
    {
        details.AppendLine();
        details.AppendLine($"{title}: {issues.Count}");
        if (issues.Count == 0)
        {
            details.AppendLine("• (yok)");
            return;
        }

        foreach (var issue in issues)
            details.AppendLine($"• {issue.Summary}");
    }

    private void ShowScrollablePopup(string content, string title, MessageBoxIcon icon)
    {
        using var dialog = new Form
        {
            Text = title,
            StartPosition = FormStartPosition.CenterParent,
            MinimizeBox = false,
            MaximizeBox = false,
            ShowInTaskbar = false,
            FormBorderStyle = FormBorderStyle.Sizable,
            ClientSize = new Size(900, 520)
        };

        var iconBox = new PictureBox
        {
            SizeMode = PictureBoxSizeMode.CenterImage,
            Width = 48,
            Dock = DockStyle.Left,
            Margin = new Padding(0, 0, 8, 0),
            Image = icon switch
            {
                MessageBoxIcon.Warning => SystemIcons.Warning.ToBitmap(),
                MessageBoxIcon.Error => SystemIcons.Error.ToBitmap(),
                _ => SystemIcons.Information.ToBitmap()
            }
        };

        var textBox = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Both,
            WordWrap = false,
            Dock = DockStyle.Fill,
            Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point),
            Text = content
        };

        var okButton = new Button
        {
            Text = "OK",
            DialogResult = DialogResult.OK,
            Anchor = AnchorStyles.Right | AnchorStyles.Bottom,
            Size = new Size(100, 32),
            Location = new Point(dialog.ClientSize.Width - 112, dialog.ClientSize.Height - 44)
        };

        var contentPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(12, 12, 12, 0)
        };
        contentPanel.Controls.Add(textBox);
        contentPanel.Controls.Add(iconBox);

        var buttonPanel = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 56,
            Padding = new Padding(12)
        };
        buttonPanel.Controls.Add(okButton);

        dialog.Controls.Add(contentPanel);
        dialog.Controls.Add(buttonPanel);
        dialog.AcceptButton = okButton;

        dialog.ShowDialog(this);
    }

    private string? PromptForExcelPassword(string fileName, bool wrongPassword)
    {
        using var dialog = new PasswordPromptForm(fileName, wrongPassword);
        return dialog.ShowDialog(this) == DialogResult.OK ? dialog.Password : null;
    }
}
