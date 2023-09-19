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

namespace xmlReader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void excel_Click(object sender, EventArgs e)
        {
            bool stop = false;
            do
            {
                string filePath = string.Empty;

                Dictionary<string, int> rules = new Dictionary<string, int>();

                rules.Add("node_name", 0);
                rules.Add("oam_enodeb_siad_oam_vlan", 0);
                rules.Add("IPV6_SIAD_OAM_IP_DEF_ROUTER", 0);
                rules.Add("IPV6_ENODEB_OAM_IP", 0);
                rules.Add("bearer_enodeb_sb_vlan_id", 0);
                rules.Add("IPV6_SIAD_BEARER_IP_DEF_ROUTER", 0);

                openFileDialog1.InitialDirectory = Application.StartupPath;
                openFileDialog1.FileName = "*.xlsx";
                openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";

                if(openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    filePath = openFileDialog1.FileName;
           
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;
                    int rowCount = xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count;

                    bool isStop = false;
                    values.Add(new List<CompanyClass>());

                    bool isRulesOk = false;
                    for (int i = 1; i <= rowCount; i++)
                    {
                        if (!isRulesOk)
                        {
                            isRulesOk = !rules.Values.Any(r => r == 0);
                        }
                        CompanyClass company = new CompanyClass();

                        List<string> cells = new List<string>();
                        for (int j = 1; j <= colCount; j++)
                        {

                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                var cellValue = xlRange.Cells[i, j].Value2.ToString();
                                if (!string.IsNullOrEmpty(cellValue))
                                {
                                    cells.Add(cellValue);
                                }
                                if (!isRulesOk)
                                {
                                    if (rules.ContainsKey(cellValue))
                                    {
                                        rules[cellValue] = j;
                                    }
                                }
                                else
                                {
                                    if (!cells.Any())
                                    {
                                        isStop = true;
                                        break;
                                    }
                                    else
                                    {

                                        if (j == rules["LINE_NO"])
                                        {
                                            company.LINE_NO = cellValue;
                                        }
                                        if (j == rules["Denk_Kabinet"])
                                        {
                                            company.Denk_Kabinet = cellValue;
                                        }
                                        if (j == rules["Kabinet_Tipi"])
                                        {
                                            company.Kabinet_Tipi = cellValue;
                                        }
                                        if (j == rules["BBU_Type"])
                                        {
                                            company.BBU_Type = cellValue;
                                        }
                                        if (j == rules["node_name"])
                                        {
                                            company.Node_name = cellValue;
                                        }

                                        if (j == rules["oam_enodeb_siad_oam_vlan"])
                                        {
                                            company.oam_enodeb_siad_oam_vlan = cellValue;
                                        }

                                        if (j == rules["IPV6_SIAD_OAM_IP_DEF_ROUTER"])
                                        {
                                            company.IPV6_SIAD_OAM_IP_DEF_ROUTER = cellValue;
                                        }

                                        if (j == rules["IPV6_ENODEB_OAM_IP"])
                                        {
                                            company.IPV6_ENODEB_OAM_IP = cellValue;
                                        }

                                        if (j == rules["ENBID"])
                                        {
                                            company.ENBID = cellValue;
                                        }
                                        if (j == rules["GNBID_Standalone"])
                                        {
                                            company.GNBID_Standalone = cellValue;
                                        }

                                        if (j == rules["bearer_enodeb_sb_vlan_id"])
                                        {
                                            company.bearer_enodeb_sb_vlan_id = cellValue;
                                        }

                                        if (j == rules["IPV6_SIAD_BEARER_IP_DEF_ROUTER"])
                                        {
                                            company.IPV6_SIAD_BEARER_IP_DEF_ROUTER = cellValue;
                                        }

                                        if (j == rules["IPV6_ENODEB_BEARER_IP"])
                                        {
                                            company.IPV6_ENODEB_BEARER_IP = cellValue;
                                        }

                                        //aLARM
                                        if (j == rules["Alarm_Definitions"])
                                        {
                                            company.Alarm_Definitions = cellValue;
                                        }

                                        if (j == rules["Normal_Contact_State"])
                                        {
                                            company.Normal_Contact_State = cellValue;
                                        }

                                        if (j == rules["Severity"])
                                        {
                                            company.Severity = cellValue;
                                        }

                                        //Uemas
                                        //1
                                        if (j == rules["EutranCellFDD_Id"])
                                        {
                                            company.EutranCellFDD_Id = cellValue;
                                        }

                                        if (j == rules["EARFCNDL"])
                                        {
                                            company.EARFCNDL = cellValue;
                                        }
                                        //2
                                        if (j == rules["EutranCellFDD_Id_2"])
                                        {
                                            company.EutranCellFDD_Id_2 = cellValue;
                                        }
                                        if (j == rules["EARFCNDL_2"])
                                        {
                                            company.EARFCNDL_2 = cellValue;
                                        }
                                        //3
                                        if (j == rules["EutranCellFDD_Id_3"])
                                        {
                                            company.EutranCellFDD_Id_3 = cellValue;
                                        }

                                        if (j == rules["EARFCNDL_3"])
                                        {
                                            company.EARFCNDL_3 = cellValue;
                                        }//4
                                        if (j == rules["EutranCellFDD_Id_4"])
                                        {
                                            company.EutranCellFDD_Id_4 = cellValue;
                                        }
                                        if (j == rules["EARFCNDL_4"])
                                        {
                                            company.EARFCNDL_4 = cellValue;
                                        }
                                        //5
                                        if (j == rules["EutranCellFDD_Id_5"])
                                        {
                                            company.EutranCellFDD_Id_5 = cellValue;
                                        }
                                        if (j == rules["EARFCNDL_5"])
                                        {
                                            company.EARFCNDL_5 = cellValue;
                                        }
                                        //6
                                        if (j == rules["EutranCellFDD_Id_6"])
                                        {
                                            company.EutranCellFDD_Id_6 = cellValue;
                                        }
                                        if (j == rules["EARFCNDL_6"])
                                        {
                                            company.EARFCNDL_6 = cellValue;
                                        }
                                        //7
                                        if (j == rules["EutranCellFDD_Id_7"])
                                        {
                                            company.EutranCellFDD_Id_7 = cellValue;
                                        }
                                        if (j == rules["EARFCNDL_7"])
                                        {
                                            company.EARFCNDL_7 = cellValue;
                                        }

                                        if (j == rules["IPV6_SIAD_BEARER_IP_DEF_ROUTER_mmbb"])
                                        {
                                            company.IPV6_SIAD_BEARER_IP_DEF_ROUTER_mmbb = cellValue;
                                        }
                                        //
                                        if (j == rules["arfcn"])
                                        {
                                            company.arfcn = cellValue;
                                        }
                                        if (j == rules["arfcn_Standalone"])
                                        {
                                            company.arfcn_Standalone = cellValue;
                                        }
                                        //For MMBB
                                        if (j == rules["Kabinet_Adi"])
                                        {
                                            company.Kabinet_Adi = cellValue;
                                        }
                                        if (j == rules["ManagedElement"])
                                        {
                                            company.ManagedElement = cellValue;
                                        }
                                        if (j == rules["node_name_secondary_id"])
                                        {
                                            company.node_name_secondary_id = cellValue;
                                        }
                                        if (j == rules["GNBID_mmbb"])
                                        {
                                            company.GNBID_mmbb = cellValue;
                                        }
                                        if (j == rules["bearer_enodeb_sb_vlan_id_mmbb"])
                                        {
                                            company.bearer_enodeb_sb_vlan_id_mmbb = cellValue;
                                        }
                                        if (j == rules["IPV6_SIAD_BEARER_IP_DEF_ROUTER_mmbb"])
                                        {
                                            company.IPV6_SIAD_BEARER_IP_DEF_ROUTER_mmbb = cellValue;
                                        }
                                        if (j == rules["IPV6_ENODEB_BEARER_IP_mmbb"])
                                        {
                                            company.IPV6_ENODEB_BEARER_IP_mmbb = cellValue;
                                        }
                                        if (j == rules["CBAND"])
                                        {
                                            company.CBAND = cellValue;
                                        }
                                        if (j == rules["CBAND_GNBID"])
                                        {
                                            company.CBAND_GNBID = cellValue;
                                        }
                                    }
                                }
                            }
                        }

                        if (isRulesOk && (!cells.Any() || cells.Count(r => !string.IsNullOrEmpty(r)) == 0))
                        {
                            break;
                        }
                        if (!string.IsNullOrEmpty(company.LINE_NO))
                        {
                            values[values.Count - 1].Add(company);
                        }
                        if (isStop)
                        {
                            break;
                        }
                    }

                    xlApp.Workbooks.Close();
                    xlApp.Quit();

                    xlRange = null;
                    xlWorksheet = null;
                    xlWorkbook = null;
                    xlApp = null;

                    sahaAdi = values[0][0].Node_name;
                    //test
                    //Dosya içerisi doluysa silme.
                    var filePath = @"C:\\PROJE DOSYALARI\\" + values[0][0].Node_name;
                    try
                    {
                        if (Directory.Exists(filePath))
                        {
                            var dir = new DirectoryInfo(filePath);
                            dir.Attributes = dir.Attributes & ~FileAttributes.ReadOnly;
                            dir.Delete(true);
                            MessageBox.Show(values[0][0].Node_name + " isimli dosya C:\\PROJE DOSYALARI\\ içerisinde bulunduğu için silindi.\nDosya oluştur butonuna tıkladığınızda yeniden üretilecektir.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    catch (IOException ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    break;
                    stop = true;
                }
                else
                {
                    MessageBox.Show("Excel dosyasını seçmediniz", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                }

            }
            while (stop == false);
        }
    }
}

