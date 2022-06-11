using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.Entity.Validation;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Security.Cryptography;
using System.Text;

namespace Excel_Upload_To_DataGridView_SQL_Server
{
    public class GlobalClass
    {
       

        public static string RemoveSingleQuote(string data)
        {
            return string.IsNullOrEmpty(data) ? string.Empty : data.Replace("'", "''");
        }

        public static string RemoveSpace(string data)
        {
            return string.IsNullOrEmpty(data) ? string.Empty : data.Trim();
        }

        public static T RemoveSingleQuote<T>(object ob)
        {
            PropertyDescriptorCollection props =
                TypeDescriptor.GetProperties(typeof(T));

            T Tob = (T)ob;
            if (Tob != null)
            {
                for (int i = 0; i < props.Count; i++)
                {
                    PropertyDescriptor prop = props[i];
                    if (prop.PropertyType.Name == "String")
                    {
                        object oPropertyValue = prop.GetValue(Tob);
                        oPropertyValue = RemoveSingleQuote((string)oPropertyValue) as object;
                        oPropertyValue = RemoveSpace((string)oPropertyValue) as object;
                        prop.SetValue(Tob, oPropertyValue);
                    }
                    //oPropertyValue
                    //table.Columns.Add(prop.Name, prop.PropertyType);
                }
            }
            return Tob;
        }

        public static string ConvertSystemDate(string date)
        {
            try
            {
                string month = date.Split('/')[1];
                string day = date.Split('/')[0];
                string year = date.Split('/')[2];

                return month + "/" + day + "/" + year;
            }
            catch (Exception ex)
            {
                return "1/1/1900";
            }
        }

        public static string GetEncryptedPassword(string password)
        {
            string pass = "";
            byte[] toEncodeAsBytes = ASCIIEncoding.ASCII.GetBytes(password);
            string returnValue = Convert.ToBase64String(toEncodeAsBytes);
            Byte[] originalBytes;
            Byte[] encodedBytes;
            MD5 objMd5 = new MD5CryptoServiceProvider();
            originalBytes = ASCIIEncoding.Default.GetBytes(returnValue);
            encodedBytes = objMd5.ComputeHash(originalBytes);
            StringBuilder ss = new StringBuilder();
            foreach (byte b in encodedBytes)
            {
                ss.Append(b.ToString("x2").ToLower());
            }
            pass = ss.ToString();
            return pass;
        }


        public static string GetMacAddress()
        {
            string macAddresses = "";

            foreach (NetworkInterface nic in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (nic.OperationalStatus == OperationalStatus.Up)
                {
                    if (nic.GetPhysicalAddress().ToString().Length > 5)
                    {
                        macAddresses = nic.GetPhysicalAddress().ToString();
                        return macAddresses;
                    }
                }
                if (nic.OperationalStatus == OperationalStatus.Down)
                {
                    if (nic.GetPhysicalAddress().ToString().Length > 5)
                    {
                        macAddresses = nic.GetPhysicalAddress().ToString();
                        return macAddresses;
                    }
                }
            }
            return macAddresses;
        }


        public static string GetMacAddressAll()
        {
            string macAddresses = "'-1'";

            foreach (NetworkInterface nic in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (nic.OperationalStatus == OperationalStatus.Up)
                {
                    if (nic.GetPhysicalAddress().ToString().Length > 5)
                    {
                        macAddresses += ",'" + nic.GetPhysicalAddress().ToString() + "'";
                        //break;
                    }
                }
                if (nic.OperationalStatus == OperationalStatus.Down)
                {
                    if (nic.GetPhysicalAddress().ToString().Length > 5)
                    {
                        macAddresses += ",'" + nic.GetPhysicalAddress().ToString() + "'";
                        //break;
                    }
                }
            }
            return macAddresses;
        }


        public void ExportData(DataTable dt, string destination)
        {
            using (var workbook = SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = workbook.AddWorkbookPart();

                workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();

                //foreach (System.Data.DataTable table in ds.Tables)
                //{

                var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                uint sheetId = 1;
                if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                {
                    sheetId =
                        sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = dt.TableName };
                sheets.Append(sheet);

                DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                List<String> columns = new List<string>();
                foreach (System.Data.DataColumn column in dt.Columns)
                {
                    columns.Add(column.ColumnName);

                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                    cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName);
                    headerRow.AppendChild(cell);
                }


                sheetData.AppendChild(headerRow);

                foreach (System.Data.DataRow dsrow in dt.Rows)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                    foreach (String col in columns)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()); //
                        newRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(newRow);
                }

                //}
            }
        }
        public DataTable ReadData(string query)
        {
            DataTable dt = new DataTable();
            SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["DBConn"].ConnectionString);
            SqlCommand cmd = new SqlCommand(query, conn);
            conn.Open();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(dt);
            conn.Close();
            da.Dispose();
            return dt;
        }



        #region crypto functions


        public static string Encrypt(string clearText)
        {
            string EncryptionKey = "MSDSL2012";
            byte[] clearBytes = Encoding.Unicode.GetBytes(clearText);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(clearBytes, 0, clearBytes.Length);
                        cs.Close();
                    }
                    clearText = Convert.ToBase64String(ms.ToArray());
                }
            }
            return clearText;
        }

        public static string Decrypt(string cipherText)
        {
            string EncryptionKey = "MSDSL2012";
            byte[] cipherBytes = Convert.FromBase64String(cipherText);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(cipherBytes, 0, cipherBytes.Length);
                        cs.Close();
                    }
                    cipherText = Encoding.Unicode.GetString(ms.ToArray());
                }
            }
            return cipherText;
        }






        public static string GetExceptionStringMsg(Exception ex)
        {

            DbEntityValidationException DbEntiyValidationException = ex as DbEntityValidationException;
            string msg = ex.Message + ex.StackTrace;

            if (DbEntiyValidationException != null)
            {
                // Retrieve the error messages as a list of strings.
                //var errorMessages = DbEntiyValidationException.EntityValidationErrors
                //        .SelectMany(x => x.ValidationErrors)
                //        .Select(x => x.ErrorMessage);
                foreach (var d in DbEntiyValidationException.EntityValidationErrors)
                {
                    foreach (var er in d.ValidationErrors)
                    {
                        msg += er.ErrorMessage;
                    }
                }


            }

            if (ex.InnerException != null)
            {
                if (ex.InnerException.InnerException != null)
                {
                    msg += Environment.NewLine + ex.InnerException.InnerException.Message;
                }
            }

            return msg;
        }


        #endregion

    }
}
