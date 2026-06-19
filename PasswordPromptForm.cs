namespace ExcelBirlestirici;

internal sealed class PasswordPromptForm : Form
{
    private readonly TextBox _txtPassword;

    public string Password => _txtPassword.Text;

    public PasswordPromptForm(string fileName, bool wrongPassword)
    {
        Text = "Excel Şifresi";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(460, 190);

        var lblMessage = new Label
        {
            AutoSize = false,
            Location = new Point(16, 16),
            Size = new Size(428, 64),
            Text = wrongPassword
                ? $"{fileName}\r\n\r\nŞifre yanlış. Lütfen tekrar deneyin:"
                : $"{fileName}\r\n\r\nBu dosya şifre korumalı. Şifreyi girin:"
        };

        var lblPassword = new Label
        {
            AutoSize = true,
            Location = new Point(16, 92),
            Text = "Şifre:"
        };

        _txtPassword = new TextBox
        {
            Location = new Point(16, 116),
            Size = new Size(428, 27),
            UseSystemPasswordChar = true,
            TabIndex = 0
        };

        var btnOk = new Button
        {
            Text = "Tamam",
            DialogResult = DialogResult.OK,
            Location = new Point(252, 150),
            Size = new Size(90, 32),
            TabIndex = 1
        };
        btnOk.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(_txtPassword.Text))
            {
                MessageBox.Show(this, "Şifre boş olamaz.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
            }
        };

        var btnCancel = new Button
        {
            Text = "İptal",
            DialogResult = DialogResult.Cancel,
            Location = new Point(354, 150),
            Size = new Size(90, 32),
            TabIndex = 2
        };

        Controls.Add(lblMessage);
        Controls.Add(lblPassword);
        Controls.Add(_txtPassword);
        Controls.Add(btnOk);
        Controls.Add(btnCancel);

        AcceptButton = btnOk;
        CancelButton = btnCancel;
    }

    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);
        _txtPassword.Focus();
    }
}
