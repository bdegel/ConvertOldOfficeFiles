using System;
using System.Windows.Forms;

namespace ConvertOldOfficeFiles
{
    public partial class FrmMain : Form
    {
        private readonly Converter _co = new Converter();

        public FrmMain()
        {
            InitializeComponent();
            Text = Application.ProductName + " Version " + Application.ProductVersion;

            _co.TextChanged += UpdateText;
            _co.StatusTextChanged += UpdateStatusText;
        }

        private void UpdateText(object? sender, EventArgs e)
        {
            tbOutput.Text = _co.Output;
        }

        private void UpdateStatusText(object? sender, EventArgs e)
        {
            statusLabel.Text = _co.StatusText;
        }

        private void btConvert_Click(object sender, EventArgs e)
        {
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                _co.ConvertPath(dlg.SelectedPath, true);
                Cursor.Current = Cursors.Default;
            }
        }

        private void btCheck_Click(object sender, EventArgs e)
        {
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                _co.ConvertPath(dlg.SelectedPath, false);
                Cursor.Current = Cursors.Default;
            }
        }

        private void FrmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            _co?.Dispose();
        }

        readonly FolderBrowserDialog dlg = new()
        {
            AutoUpgradeEnabled = false,
            ShowNewFolderButton = false,
            RootFolder = Environment.SpecialFolder.Desktop,
            Description = "Select directory to be converted",
            UseDescriptionForTitle = false
        };
    }
}