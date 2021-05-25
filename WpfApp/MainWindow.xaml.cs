using System.Collections.Generic;
using System.Windows;
using System.Windows.Forms;
using SW;
using MessageBox = System.Windows.Forms.MessageBox;
using OpenFileDialog = System.Windows.Forms.OpenFileDialog;

namespace WpfApp
{
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private string[] solidFiles;
        private string drawingPath;
        
        private void SolidFiles_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = @"P:\1. BEIN";
            //openFileDialog.InitialDirectory = @"S:\Programs\Macros\SW TreeManager\19-816-04 - Additional Items";
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "SolidWorks Drawings (*.slddrw)|*.slddrw";
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                solidFiles = openFileDialog.FileNames;
            }
        }

        private void GetDirectory_Click(object sender, RoutedEventArgs e)
        {
            var folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.SelectedPath = @"P:\1. BEIN";
            //folderBrowserDialog.SelectedPath = @"S:\Programs\Macros\SW TreeManager\19-816-04 - Additional Items\Drawings";
            folderBrowserDialog.ShowDialog();
            drawingPath = folderBrowserDialog.SelectedPath + @"\";
        }

        private void Ready_Click(object sender, RoutedEventArgs e)
        {
            if(solidFiles == null)
            {
                MessageBox.Show("Choose SolidWorks drawings first!");
            }
            else if (drawingPath == null)
            {
                MessageBox.Show("Choose PDF drawings directory first!");
            }
            else
            {
                List<string> props = new List<string>();
                props.Add(Input01.Text);
                props.Add(Input02.Text);
                props.Add(Input03.Text);
                props.Add(Input04.Text);
                props.Add(Input05.Text);
                props.Add(Input06.Text);
                props.Add(Input07.Text);
                props.Add(Input08.Text);
                props.Add(Input09.Text);
                props.Add(Input10.Text);
                SwTools swTools = new SwTools();
                swTools.GetFiles(solidFiles, drawingPath, props);
            }
        }
    }
}