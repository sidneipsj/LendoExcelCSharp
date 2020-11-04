using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //openFileDialog1.Filter = "*.xls | Microsoft Excel";


            openFileDialog1.Title = "Selecione o Arquivo";


            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                DataSet ds = new DataSet();

                OleDbConnection conexao = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" +

                "Data Source=" + openFileDialog1.FileName + ";" +

                "Extended Properties=Excel 8.0;");
                try
                {
                    OleDbDataAdapter da = new OleDbDataAdapter("Select * From [Plan1$]", conexao);
                    da.Fill(ds);
                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);
                }

                vGrade.DataSource = ds.Tables[0];

                vGrade.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCellsExceptHeader;

                conexao.Close();

            }




        }
    }
}
