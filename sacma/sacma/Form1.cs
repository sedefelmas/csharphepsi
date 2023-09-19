using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace sacma
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string cellValue = null;
        bool stop = false;
        List<List<CompanyClass>> values = new List<List<CompanyClass>>();
        private void button1_Click(object sender, EventArgs e)
        {
            do
            {
                string filePath = string.Empty;
                openFileDialog1.InitialDirectory = Application.StartupPath;
                openFileDialog1.FileName = "*.xlsx";
                openFileDialog1.Filter = "xlsx files (*.xlsx]|*.xlsx|All files (*.*]|*.*";                
                if(openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    Dictionary<string, int> rules = new Dictionary<string, int>();
                    rules.Add("node_name", 0);
                    rules.Add("oam_enodeb_siad_oam_vlan", 0);
                    rules.Add("IPV6_SIAD_OAM_IP_DEF_ROUTER", 0);
                    rules.Add("IPV6_ENODEB_OAM_IP", 0);
                    rules.Add("bearer_enodeb_sb_vlan_id", 0);
                    rules.Add("IPV6_SIAD_BEARER_IP_DEF_ROURER", 0);
                    filePath = openFileDialog1.FileName;
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;
                    int rowCount = xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count;
                    values.Add(new List<CompanyClass>());
                    for(int i = 1; i <= rowCount; i++)
                    {
                        CompanyClass company = new CompanyClass();
                        List<string> cells = new List<string>();
                        for(int j = 1; j <= colCount; j++)
                        {
                            if(xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                cellValue = xlRange.Cells[i, j].Value2.ToString();

                            }
                        }
                    }

                }
            }
            while (stop == false);
        }
    }
}
