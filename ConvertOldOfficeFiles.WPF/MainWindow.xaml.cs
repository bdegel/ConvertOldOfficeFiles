using Ookii.Dialogs.Wpf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ConvertOldOfficeFiles.WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly Converter _co = new();

        public MainWindow()
        {
            InitializeComponent();
            Title = Application.Current.MainWindow.GetType().Assembly.GetName().Name + " Version " + System.Reflection.Assembly.GetEntryAssembly().GetName().Version;

            _co.TextChanged += UpdateText;
            _co.StatusTextChanged += UpdateStatusText;
        }

        private void UpdateText(object? sender, EventArgs e)
        {
            tbOutput.Text = _co.Output;
        }

        private void UpdateStatusText(object? sender, EventArgs e)
        {
            statusLabel.Content = _co.StatusText;
        }

        private void btConvert_Click(object sender, RoutedEventArgs e)
        {
            if (_dlg.ShowDialog() != false) return;

            Mouse.OverrideCursor = Cursors.Wait;
            _co.ConvertPath(_dlg.SelectedPath, true);
            Mouse.OverrideCursor = null;
        }

        private void btCheck_Click(object sender, RoutedEventArgs e)
        {
            if (_dlg.ShowDialog() != true) return;

            Mouse.OverrideCursor = Cursors.Wait;
            _co.ConvertPath(_dlg.SelectedPath, false);
            Mouse.OverrideCursor = null;
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            _co?.Dispose();
        }

        readonly VistaFolderBrowserDialog _dlg = new()
        {
            ShowNewFolderButton = false,
            RootFolder = Environment.SpecialFolder.Desktop,
            Description = "Select directory to be converted",
            UseDescriptionForTitle = false
        };
    }
}
