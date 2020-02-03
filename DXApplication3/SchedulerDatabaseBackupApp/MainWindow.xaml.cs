using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using DevExpress.Xpf.Core;
using Microsoft.Win32;

namespace DatabaseBackupApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : ThemedWindow
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        public string GetFilePath(string initialDirectory = "")
        {
            if (initialDirectory == "")
            {
                initialDirectory = @"\\s-fs1-smdrv\mydocs$\Joshua.Meservey\Microsoft Access";
            }

            string filename;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = initialDirectory;
            openFileDialog.Filter = "Excel Files (*.accdb)|*.accdb";
            Nullable<bool> result = Convert.ToBoolean(openFileDialog.ShowDialog());

            filename = openFileDialog?.FileName;

            return filename;
        }

        private void CopyDBButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Database.CopyDatabase(DestinationDBTextBox.Text);
                MessageBox.Show("Finished!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + "\n\n" + ex.StackTrace);
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);
            }
        }

        private void ChooseDestinationDBButton_Click(object sender, RoutedEventArgs e)
        {
            DestinationDBTextBox.Text = GetFilePath();
        }

        private void ChooseSourceDBButton_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
