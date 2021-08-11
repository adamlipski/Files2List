using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;

namespace Files2List
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string filePath { get; set; }
        public string folderPath { get; set; }
        string [] AllLines;

        List<string> AllFilesNames { get; set; }

   
        public MainWindow()
        {
            InitializeComponent();

            Excel excel = new Excel();
        }

        private void openFolder_Click(object sender, RoutedEventArgs e)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet worksheet = workbook.CreateSheet("Sheet1");
            OpenFileDialog openAnyFile = new OpenFileDialog();

            if(openAnyFile.ShowDialog() == true)
            {
                filePath = openAnyFile.FileName;
                string fileName = openAnyFile.SafeFileName;
                //var filesNames = openAnyFile.SafeFileNames;

                folderPath = filePath.Remove(folderPath.Count() - fileName.Count());

                AllLines = Directory.GetFiles(folderPath);
                AllFilesNames = AllLines.ToList();

                int rownum = 0;
                int cellnum = 0;
                foreach (string line in AllFilesNames)
                {

                    IRow row = worksheet.CreateRow(rownum);
                    ICell cell = row.CreateCell(cellnum);
                    cell.SetCellValue(line.Substring(folderPath.Count()));
                    rownum++;
                }

                FileStream newWorkBook = File.Create($"{folderPath}ListOfFiles.xlsx");
                workbook.Write(newWorkBook);
                newWorkBook.Close();


            }
            


        }
    }
}
