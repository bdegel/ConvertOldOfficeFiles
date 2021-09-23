using System;
using System.IO;
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
        }

        private void UpdateText(object? sender, EventArgs e)
        {
            tbOutput.Text = _co.Output;
        }

        private void btConvert_Click(object sender, EventArgs e)
        {   
            tbOutput.Clear();
            _co.ConvertPath(dlg.SelectedPath, true);
            statusLabel.Text = "Ready";
            Cursor.Current = Cursors.Default;
            tbOutput.AppendText(_co.FileCount + " files converted" + Environment.NewLine);
        }

        private void btCheck_Click(object sender, EventArgs e)
        {
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                tbOutput.Clear();
                _co.ConvertPath(dlg.SelectedPath, false);
                statusLabel.Text = "Ready";
                Cursor.Current = Cursors.Default;
                tbOutput.AppendText(_co.FileCount + " files found" + Environment.NewLine);

                btConvert.Enabled = true;
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