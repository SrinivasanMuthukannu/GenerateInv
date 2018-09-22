using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Invoicer.GUI;

namespace InvoiceGUI
{
    public partial class Form1 : Form
    {
        DataTable dtexcel = new DataTable();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            {
                filePath = file.FileName; //get the path of the file  
                fileExt = Path.GetExtension(filePath); //get the file extension  
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        DataTable dtExcel = new DataTable();
                        dtExcel = ReadExcel(filePath, fileExt); //read excel file  
                        dataGridView1.Visible = true;
                        dataGridView1.DataSource = dtExcel;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                }
            }  
        }

        public DataTable ReadExcel(string fileName, string fileExt)
        {
            DataTable dt = new DataTable();
            string conn = string.Empty;
            
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dt); //fill excel data into dataTable  
                    IEnumerable<DataRow> newRows = dt.AsEnumerable().Skip(2);
                    dtexcel = newRows.CopyToDataTable();
                }
                catch { }
            }
            return dtexcel;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DataView dv = new DataView(dtexcel);
            dv.RowFilter = "F3"+ " LIKE '%" + BillNumber.Text + "%'";

            //BindingSource bs = new BindingSource();
            //bs.DataSource = dataGridView1.DataSource;
            //bs.Filter = dataGridView1.Columns[2].HeaderText.ToString() + " LIKE '%" + BillNumber.Text + "%'";
            dataGridView1.DataSource = dv;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DataView dv = new DataView(dtexcel);
            dv.RowFilter = "F3" + " LIKE '%" + BillNumber.Text + "%'";
            DataTable dt = new DataTable();
            dt = dv.ToTable();
            new Process().Go(dt, BillNumber.Text);
        }

       
    }
}
