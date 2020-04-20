
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Configuration;
using System.Data.SqlClient;


namespace ExportToSQL
{
    public partial class Form1 : Form
    {
        //chosen connection string will be based on the excel extension, either xls or xlsx
        private string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }



        private void Selectbtn_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            //getting the file path 
            string filePath = openFileDialog1.FileName;
            string extension = Path.GetExtension(filePath);
            string conString = "";
            string sheetName = "";

            //choosing type of connection string
            switch (extension)
            {
                case ".xls":
                    conString = string.Format(Excel03ConString, filePath, "YES");
                    break;
                case ".xlsx":
                    conString = string.Format(Excel07ConString, filePath, "YES");
                    break;
            }
            //creating excel connection string 
            using (OleDbConnection conn = new OleDbConnection(conString))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    cmd.Connection = conn;
                    conn.Open();
                    //getting the sheet name
                    DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    sheetName = dt.Rows[0]["Table_Name"].ToString();
                    conn.Close();
                }
            }

            //reading the sheet data into a DataTable
            using (OleDbConnection con = new OleDbConnection(conString))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    OleDbDataAdapter oda = new OleDbDataAdapter();
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = con;
                    con.Open();
                    oda.SelectCommand = cmd;
                    DataTable dt = new DataTable();
                    oda.Fill(dt);
                    con.Close();
                    //binding the DataTable to DataGridView control
                    dataGridView1.DataSource = dt;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //creating a DataTable making the columns same as that of SQL table to which it'll be saved 
            DataTable dt = new DataTable();
            dt.Columns.AddRange(new DataColumn[11]{
               new DataColumn("PaymentDate", typeof(string)), /* could have to be changed*/
                new DataColumn("Name", typeof(string)),
                new DataColumn("SortCode", typeof(int)),
                new DataColumn("Narration", typeof(string)),
                new DataColumn("Amount", typeof(double)),
                new DataColumn("ValueDate", typeof(string)), /* could have to be changed*/
                new DataColumn("TranID", typeof(string)),
                new DataColumn("Period", typeof(string)),
                new DataColumn("Count", typeof(string)),
                new DataColumn("Type", typeof(string)),
                new DataColumn("PaymentNarration", typeof(string)),
            });

            //populating the dataTable rows with datagrid
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {

                string paymentDate = row.Cells[0].Value.ToString();
                string name = row.Cells[1].Value.ToString();
                int sortCode = Convert.ToInt32(row.Cells[2].Value);
                string narration = row.Cells[3].Value.ToString();
                double amount = Convert.ToDouble(row.Cells[4].Value);
                string valueDate = row.Cells[5].Value.ToString(); /* could have to be changed*/
                string tranID = row.Cells[6].Value.ToString();
                string period = row.Cells[7].Value.ToString();
                string count = row.Cells[8].Value.ToString();
                string type = row.Cells[9].Value.ToString();
                string paymentNarration = row.Cells[10].Value.ToString();
                dt.Rows.Add(paymentDate, name, sortCode, narration, amount, valueDate, tranID, period, count, type, paymentNarration);
            }
            //using sqlbulkcopy to copy data to database
            if (dt.Rows.Count > 0)
            {
                string constring = @"Data Source=AISHATAKINYEMI\SQLEXPRESS;Initial Catalog=db.mdl;Integrated Security=True";
                using (SqlConnection con = new SqlConnection(constring))
                {
                    using (SqlBulkCopy exToSql = new SqlBulkCopy(con))
                    {
                        //mapping columns of DataTable to sql database columns
                        exToSql.DestinationTableName = "Temp";
                        exToSql.ColumnMappings.Add(0, 1);
                        exToSql.ColumnMappings.Add(1, 2);
                        exToSql.ColumnMappings.Add(2, 3);
                        exToSql.ColumnMappings.Add(3, 4);
                        exToSql.ColumnMappings.Add(4, 5);
                        exToSql.ColumnMappings.Add(5, 6);
                        exToSql.ColumnMappings.Add(6, 7);
                        exToSql.ColumnMappings.Add(7, 8);
                        exToSql.ColumnMappings.Add(8, 9);
                        exToSql.ColumnMappings.Add(9, 10);
                        exToSql.ColumnMappings.Add(10, 11);
                        con.Open();
                        exToSql.WriteToServer(dt);
                        con.Close();
                    }
                }
            }
        }
    }


        }