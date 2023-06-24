using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;

namespace ParamerticModuleModelling
{


    public partial class InputParameter_stair : System.Windows.Forms.Form
    {
        public static string ST_name;
        public static string Modelselect;

        public static string AA,AB,AC,AD,AE,AF,AG;
        public static string BA, BB;
        public static string CA, CB;
        public static string DA, DB, DC, DD, DE, DF, DG;

        public InputParameter_stair()
        {

            InitializeComponent();



            cmb_stairtype.Items.Add("板式梯梁钢楼梯系统");
            cmb_stairtype.Items.Add("型钢梯梁钢楼梯系统");
            cmb_stairtype.Items.Add("槽钢梯梁钢楼梯系统");

            cmb_stairtype.SelectedIndex = 0;




            string f_path = @" Y:\数字化课题\族库\楼梯系统模块";


            if (Directory.Exists(f_path) == false)
            {
                f_path = @"C:\ProgramData\Autodesk\Revit\Addins\2018\族库\楼梯系统模块";
            }




            List<string> fileNames = Get_rfa_FileNames(f_path );
            List<string> fileNames_filter =new List<string>();

            // 获取选择项中的前两个字符
            string selectedPrefix = cmb_stairtype.SelectedItem.ToString().Substring(0, 2);

            foreach (string file in fileNames)
            {
                if (file.Substring(0,2) == (selectedPrefix) && file.Contains("双跑"))
                {
                    fileNames_filter.Add(file);
                }
            }
            cmb_ModelSelect.DataSource = fileNames_filter;

            //cmb_ModelSelect.SelectedIndex = 0;
        }
        private List<string> Get_rfa_FileNames(string folderPath)
        {
            List<string> fileNames = new List<string>();

            // 获取当前文件夹下所有后缀为.rfa的文件
            string[] files = Directory.GetFiles(folderPath, "*.rfa", SearchOption.AllDirectories);

            foreach (string file in files)
            {
                // 获取文件名（不含扩展名和路径）
                string fileName = Path.GetFileNameWithoutExtension(file);
                fileNames.Add(fileName);
            }

            return fileNames;
        }

        public void ParameterTransfer()
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void Textbox_BA_TextChanged(object sender, EventArgs e)
        {

        }

        private void Textbox_BB_TextChanged(object sender, EventArgs e)
        {

        }

        private void Textbox_CA_TextChanged(object sender, EventArgs e)
        {

        }

        private void Textbox_CB_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string f_path = @" Y:\数字化课题\族库\楼梯系统模块";


            if (Directory.Exists(f_path) == false)
            {
                f_path = @"C:\ProgramData\Autodesk\Revit\Addins\2018\族库\楼梯系统模块";
            }

            List<string> fileNames = Get_rfa_FileNames(f_path);
            List<string> fileNames_filter = new List<string>();

            // 获取选择项中的前两个字符
            string selectedPrefix = cmb_stairtype.SelectedItem.ToString().Substring(0, 2);

            foreach (string file in fileNames)
            {
                if (file.Substring(0, 2)==(selectedPrefix)&& file.Contains("双跑"))
                {
                    fileNames_filter.Add(file);
                }
            }
            cmb_ModelSelect.DataSource = fileNames_filter;

            //cmb_ModelSelect.SelectedIndex = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            ST_name = ST_label.Text;
            Modelselect = cmb_ModelSelect.SelectedItem.ToString();

            AA = Textbox_AA.Text;
            AB = Textbox_AB.Text;
            AC = Textbox_AC.Text;
            AD = Textbox_AD.Text;
            AE = Textbox_AE.Text;
            AF = Textbox_AF.Text;
            AG = Textbox_AG.Text;

            BA = Textbox_BA.Text;
            BB = Textbox_BB.Text;
            
            CA = Textbox_CA.Text;
            CB = Textbox_CB.Text;

            DA = Textbox_DA.Text;
            DB = Textbox_DB.Text;
            DC = Textbox_DC.Text;
            DD = Textbox_DD.Text;
            DE = Textbox_DE.Text;
            DF = Textbox_DF.Text;
            DG = Textbox_DG.Text;


            this.Close();
        }
    }
}
