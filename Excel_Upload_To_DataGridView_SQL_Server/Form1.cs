
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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

      

       

        private void txt_Url_Click(object sender, EventArgs e)
        {
            //https://www.c-sharpcorner.com/UploadFile/bd6c67/how-to-create-excel-file-using-C-Sharp/
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = @"C:\";
            saveFileDialog1.Title = "Save Excel Files";

            saveFileDialog1.DefaultExt = "xlsx";
            saveFileDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            saveFileDialog1.FilterIndex = 1;
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txt_Url.Text = saveFileDialog1.FileName;

            }
            else
            {
                return;
            }
        }

        private void btn_format_Click(object sender, EventArgs e)
        {
            //https://www.c-sharpcorner.com/UploadFile/bd6c67/how-to-create-excel-file-using-C-Sharp/
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            //Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = "EmployeeSheet";

                //worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[1, 8]].Merge();
                worKsheeT.Cells[1, 1] = "SL_NO";
                worKsheeT.Cells[1, 2] = "EmployeeName";
                worKsheeT.Cells[1, 3] = "Phone";
                worKsheeT.Cells[1, 4] = "Email";
                worKsheeT.Cells[1, 5] = "Address";
                worKsheeT.Cells.Font.Size = 15;

                //int rowcount = 1;              
                //celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[rowcount, ExportToExcel().Columns.Count]];
                //celLrangE.EntireColumn.AutoFit();
                //Microsoft.Office.Interop.Excel.Borders border = celLrangE.Borders;
                //border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                //border.Weight = 2d;


                //celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[2, ExportToExcel().Columns.Count]];


                worKbooK.SaveAs(txt_Url.Text); ;
                worKbooK.Close();
                excel.Quit();
                MessageBox.Show("Successfully Create Excel File");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
            finally
            {
                worKsheeT = null;
                worKbooK = null;
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
    }
}
