
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;

namespace Excel_Upload_To_DataGridView_SQL_Server
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }        

        private void btn_format_Click(object sender, EventArgs e)
        {
            try
            {
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Columns.Add("SL_NO", typeof(int));
                dt.Columns.Add("EmployeeName", typeof(string));
                dt.Columns.Add("Phone", typeof(string));
                dt.Columns.Add("Email", typeof(string));
                dt.Columns.Add("Address", typeof(string));                
                if (dt == null)
                {
                    MessageBox.Show("No item found");
                    return;
                }
                FolderBrowserDialog fbd = new FolderBrowserDialog();
                DialogResult result = fbd.ShowDialog();
                if (!string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    dt.TableName = "EmployeeSheet";
                    string fileName = "\\EmployeeList" + DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString() + ".xls";
                    GlobalClass gc = new GlobalClass();
                    gc.ExportData(dt, fbd.SelectedPath + fileName);
                }
                Cursor.Current = Cursors.Default;
                MessageBox.Show("Save Successfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void btn_ExcelUpload_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file
            if (file.ShowDialog() == DialogResult.OK) //if there is a file chosen by the user
            {
                string fileExt = Path.GetExtension(file.FileName); //get the file extension
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        DataTable dtExcel = ReadExcel(file.FileName); //read excel file
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

        private DataTable GetDataTable(string sql, string connectionString)
        {
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                System.Data.OleDb.OleDbDataAdapter adapter = new System.Data.OleDb.OleDbDataAdapter(sql, conn);
                System.Data.DataSet ds = new System.Data.DataSet();
                adapter.Fill(ds);
                return ds.Tables[0];
            }
        }
        private DataTable ReadExcel(string fileName)
        {
            // if error -> The 'Microsoft.ACE.OLEDB.12.0' provider is not registered on the local machine
            // in this menu: project -> yourproject properties... -> Build : uncheck "prefer 32-Bit"
            string connString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=yes'", fileName);
            DataTable dt = GetDataTable("SELECT * from [EmployeeSheet$]", connString);
            return dt;
        }

        private void btn_save_Click(object sender, EventArgs e)
        {
            try
            {
                string connectionString = ConfigurationManager.ConnectionStrings["DBConn"].ConnectionString;
                using (SqlConnection conn = new SqlConnection(connectionString))
                {

                    conn.Open();
                    int count = 0;
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {

                        if (!row.IsNewRow)
                        {
                            using (SqlCommand cmd = new SqlCommand(" INSERT INTO [dbo].[Employee]([EmpName],[Phone],[Email],[Address]) VALUES(@c1,@c2,@c3,@c4)", conn))
                            {

                                cmd.Parameters.AddWithValue("@C1", row.Cells[1].Value);
                                cmd.Parameters.AddWithValue("@C2", row.Cells[2].Value);
                                cmd.Parameters.AddWithValue("@C3", row.Cells[3].Value);
                                cmd.Parameters.AddWithValue("@C4", row.Cells[4].Value);

                                cmd.ExecuteNonQuery();
                                count++;
                            }
                            
                        }                       
                    }
                    conn.Close();
                    MessageBox.Show(count + " Rows inserted.");
                    dataGridView1.DataSource = null;
                    dataGridView1.Rows.Clear();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);              
            }
            
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {

                
                GlobalClass gc = new GlobalClass();
                System.Data.DataTable dt = new System.Data.DataTable();

                dt = gc.ReadData("SELECT * FROM EMPLOYEE");
                if (dt == null)
                {
                    MessageBox.Show("No item found");
                    return;
                }
                FolderBrowserDialog fbd = new FolderBrowserDialog();
                DialogResult result = fbd.ShowDialog();
                if (!string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    dt.TableName = "EmployeeSheet";
                    string fileName = "\\EmployeeList" + DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString() + ".xls";                    
                    gc.ExportData(dt, fbd.SelectedPath + fileName);
                }
                Cursor.Current = Cursors.Default;
                MessageBox.Show("Save Successfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }


        }
    }
}
