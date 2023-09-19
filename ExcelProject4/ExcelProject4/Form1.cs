using System;
using OfficeOpenXml;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace ExcelProject4
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;

            openFileDialog1.InitialDirectory = Application.StartupPath;
            openFileDialog1.FileName = "*.xlsx";
            openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";


            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog1.FileName;

                FileInfo fileInfo = new FileInfo(filePath);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
                {
                    ExcelWorkbook excelWorkbook = excelPackage.Workbook;
                    ExcelWorksheet excelWorksheet = excelWorkbook.Worksheets[1];
                    int colCount = excelWorksheet.Dimension.End.Column;
                    int rowCount = excelWorksheet.Dimension.End.Row;
                    for (int i = 1; i <= rowCount; i++)
                    {
                        for (int j = 0; j < colCount; j++)
                        {
                            listBox1.Items.Add("Row: " + " Column:" + j + " Value: " + excelWorksheet.Cells[i, j].Value.ToString());

                        }
                    }
                }
            }
        }
    }
}
