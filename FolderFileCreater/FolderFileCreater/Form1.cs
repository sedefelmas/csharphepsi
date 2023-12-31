﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace FolderFileCreater
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string mosCellValue = null;
        string xmlCellValue = null;
        private void button1_Click(object sender, EventArgs e)
        {
            string cellValue = null;
            bool stop = false;
            string excelFileName = null;
            int index = 0;
            do
            {
                string excelFilePath = string.Empty;
                openFileDialog1.InitialDirectory = Application.StartupPath;
                openFileDialog1.FileName = "*.xlsx";
                openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    excelFilePath = openFileDialog1.FileName;
                    index = excelFilePath.LastIndexOf("\\") + 1;
                    excelFileName = excelFilePath.Substring(index);
                    MessageBox.Show(excelFileName + " dosyası seçildi.");


                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(excelFilePath);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;
                    int rowCount = xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count;
                    string[][] lists = new string[rowCount][];
                    for (int i = 0; i <= rowCount - 1; i++)
                    {
                        lists[i] = new string[colCount];
                        for (int j = 0; j <= colCount - 1; j++)
                        {
                            if (xlRange.Cells[i + 1, j + 1] != null && xlRange.Cells[i + 1, j + 1].Value2 != null)
                            {
                                cellValue = xlRange.Cells[i + 1, j + 1].Value2.ToString();
                                lists[i][j] = cellValue;
                            }
                        }
                    }
                    for (int i = 0; i < lists.Length; i++)
                    {
                        for (int j = 0; j < lists[i].Length; j++)
                        {
                            if (lists[i][j] == "node_name")
                            {
                                mosCellValue = lists[i + 1][j];
                            }
                            else if (lists[i][j] == "IPV6_SIAD_BEARER_IP_DEF_ROUTER")
                            {
                                xmlCellValue = lists[i + 1][j];
                            }
                            if (mosCellValue != null && xmlCellValue != null)
                            {
                                break;
                            }
                        }
                    }
                    button4.Enabled = true;
                    xlApp.Workbooks.Close();
                    xlApp.Quit();

                    xlRange = null;
                    xlWorksheet = null;
                    xlWorkbook = null;
                    xlApp = null;
                    stop = true;
                }
                else
                {
                    MessageBox.Show("Excel dosyasını seçmediniz!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                }
            }
            while (stop == false);

        }
        string xmlFileName = string.Empty;
        string xmlFilePath = string.Empty;
        private void button2_Click(object sender, EventArgs e)
        {
            string xmlFilePath = string.Empty;
            openFileDialog2.InitialDirectory = Application.StartupPath;
            openFileDialog2.FileName = "*.xml";
            openFileDialog2.Filter = "xml files (*.xml)|*.xml|All files (*.*)|*.*";
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                string xmlText = "";
                List<string> xmlLines = new List<string>();
                xmlFilePath = openFileDialog2.FileName;
                StreamReader sr = new StreamReader(xmlFilePath);
                while (sr.EndOfStream == false)
                {
                    xmlLines.Add(sr.ReadLine());
                }
                sr.Close();
                foreach (string line in xmlLines)
                {
                    xmlText += line + "\r\n";
                }
                if (xmlText.Contains("Yupana_Ip"))
                {
                    xmlText = xmlText.Replace("Yupana_Ip", xmlCellValue);
                    xmlFileName = "00_RbsSummary.xml";
                    xmlFilePath = Path.Combine(nodeFolderPath, xmlFileName);
                    StreamWriter sw = new StreamWriter(xmlFilePath);
                    sw.WriteLine(xmlText);
                    sw.Close();
                }
                else
                {
                    MessageBox.Show("Dosya kaydedilemedi!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


            }
        }
        string mosFileName = string.Empty;
        string mosFilePath = string.Empty;
        private void button3_Click(object sender, EventArgs e)
        {
            string mosText = "";
            List<string> mosLines = new List<string>();
            string mosFilePath = string.Empty;
            openFileDialog3.InitialDirectory = Application.StartupPath;
            openFileDialog3.FileName = "*.mos";
            openFileDialog3.Filter = "mos files (*.mos)|*.mos|All files (*.*)|*.*";
            if (openFileDialog3.ShowDialog() == DialogResult.OK)
            {
                mosFilePath = openFileDialog3.FileName;
                StreamReader sr = new StreamReader(mosFilePath);
                while (sr.EndOfStream == false)
                {
                    mosLines.Add(sr.ReadLine());
                }
                sr.Close();
                foreach (string line in mosLines)
                {
                    mosText += line + "\r\n";
                }
                if (mosText.Contains(" Yupana"))
                {
                    mosCellValue = " " + mosCellValue;
                    mosText = mosText.Replace(" Yupana", mosCellValue);
                    mosFileName = "01_SiteBasic.mos";
                    mosFilePath = Path.Combine(nodeFolderPath, mosFileName);
                    StreamWriter sw = new StreamWriter(mosFilePath);
                    sw.WriteLine(mosText);
                    sw.Close();
                    //File.WriteAllText(mosFilePath, mosText);
                }
                else
                {
                    MessageBox.Show("Dosya kaydedilemedi!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        string projectFolderPath = string.Empty;
        string nodeFolderPath = string.Empty;


        private void button4_Click(object sender, EventArgs e)
        {
            projectFolderPath = "C:\\Proje Dosyaları\\";
            nodeFolderPath = projectFolderPath + mosCellValue;
            Directory.CreateDirectory(projectFolderPath);
            Directory.CreateDirectory(nodeFolderPath);
            button2.Enabled = true;
            button3.Enabled = true;
        }
    }
}
