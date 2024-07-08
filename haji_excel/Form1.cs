using System;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using OfficeOpenXml;

namespace haji_excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();
        }
        public string[] ReadColumnA(string filePath)
        {
            foreach (var item in listBox1.Items)
            {
                if (!File.Exists(item.ToString()))
                {
                    MessageBox.Show($"فایل {item.ToString()} یافت نشد");
                }
            }
            List<string> columnAValues = new List<string>();

            FileInfo fileInfo = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // اولین شیت

                int rows = worksheet.Dimension.Rows;

                for (int i = 2; i <= rows; i++)
                {
                    string cellValue = worksheet.Cells[i, 2].Text; // Column B
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        columnAValues.Add(cellValue);
                    }
                }
                for (int i = 2; i <= rows; i++)
                {
                    string cellValue = worksheet.Cells[i, 3].Text; // Column B
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        columnAValues.Add(cellValue);
                    }
                }
            }
            HashSet<string> uniqueValues = new HashSet<string>(columnAValues);
            return uniqueValues.ToArray(); 
        }
        private void button2_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                openFileDialog.Title = "Select an Excel File";
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    foreach (string filename in openFileDialog.FileNames)
                    {
                        listBox1.Items.Add(filename);
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<string> filePaths = new List<string>();
            foreach (object item in listBox1.Items)
            {
                string filePath = item.ToString();
                filePaths.Add(filePath);
            }
            string[] filePathsArray = filePaths.ToArray();
            string[][] columnAValuesArray = new string[filePaths.Count][];
            for (int i = 0; i < filePaths.Count; i++)
            {
                columnAValuesArray[i] = ReadColumnA(filePaths[i]);
            }
            for (int i = 0; i < columnAValuesArray.Length; i++)
            {
                for (int j = i + 1; j < columnAValuesArray.Length; j++)
                {
                    for (int k = 0; k < columnAValuesArray[i].Length; k++)
                    {
                        for (int l = 0; l < columnAValuesArray[j].Length; l++)
                        {
                            if (columnAValuesArray[i][k] == columnAValuesArray[j][l])
                            {
                                textBox1.AppendText($"{columnAValuesArray[i][k]} >> [{Path.GetFileName(filePaths[i])} | {Path.GetFileName(filePaths[j])}]");
                                textBox1.AppendText(Environment.NewLine); 
                            }
                        }
                    }
                }
            }
            MessageBox.Show("finish");
        }
        private void button3_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
        }
    }
}