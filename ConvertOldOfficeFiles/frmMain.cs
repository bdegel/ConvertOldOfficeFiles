using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ConvertOldOfficeFiles
{
    public partial class FrmMain : Form
    {
        public FrmMain()
        {
            InitializeComponent();
            Text = Application.ProductName + " Version " + Application.ProductVersion;

            Converter.TextChanged += UpdateText;
        }

        private void UpdateText(object? sender, EventArgs e)
        {
            tbOutput.Text = Converter.Output;
        }

        private void btConvert_Click(object sender, EventArgs e)
        {
            var path = tbPath.Text.Trim();
            if (path.Length <= 0 || !Directory.Exists(path)) return;
            
            tbOutput.Clear();
            Converter.ConvertPath(path, true);
            statusLabel.Text = "Ready";
            Cursor.Current = Cursors.Default;
            tbOutput.AppendText(Converter.FileCount + " files converted" + Environment.NewLine);
        }

        private void btCheck_Click(object sender, EventArgs e)
        {
            var path = tbPath.Text.Trim();
            if (path.Length <= 0 || !Directory.Exists(path)) return;
            
            tbOutput.Clear();
            Converter.ConvertPath(path, false);
            statusLabel.Text = "Ready";
            Cursor.Current = Cursors.Default;
            tbOutput.AppendText(Converter.FileCount + " files found" + Environment.NewLine);
        }

        private void frmMain_Shown(object sender, EventArgs e)
        {
            tbPath.Focus();
        }

        private void btSelectPath_Click(object sender, EventArgs e)
        {
            var dlg = new FolderBrowserDialog
            {
                AutoUpgradeEnabled = false,
                ShowNewFolderButton = false,
                RootFolder = Environment.SpecialFolder.Desktop,
                Description = "Select directory to be converted",
                UseDescriptionForTitle = false
            };

            if (dlg.ShowDialog() == DialogResult.OK)
                tbPath.Text = dlg.SelectedPath;
        }
    }
}