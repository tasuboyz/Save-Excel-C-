using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Interfaccia_Database_sql_C_
{
    public class Table
    {
        //string connectionString = "Data Source=RITSRVVBI01;Initial Catalog=DW;User ID=interfacce;Password=rub";       
        string ExcelPath;
        public DataTable ReadExcel()
        {
            DirectoryPaths paths = new DirectoryPaths();
            var ExcelFilePath = paths.GetExcelIniPath();
            IniFile iniFile = new IniFile(ExcelFilePath);
            var dataTable = new DataTable();
            if (File.Exists(ExcelFilePath))
            {
                ExcelPath = iniFile.Read("Settings", "ExcelPath");
                if (!string.IsNullOrEmpty(ExcelPath))
                {
                    try
                    {
                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                        using (var package = new ExcelPackage(new FileInfo(ExcelPath)))
                        {
                            var worksheet = package.Workbook.Worksheets[0];
                            if (worksheet != null)
                            {
                                var hasHeader = true;
                                foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                                {
                                    dataTable.Columns.Add(hasHeader ? firstRowCell.Text : $"Column {firstRowCell.Start.Column}");
                                }
                                var startRow = hasHeader ? 2 : 1;
                                for (var rowNum = startRow; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                                {
                                    var wsRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
                                    var row = dataTable.NewRow();
                                    foreach (var cell in wsRow)
                                    {
                                        row[cell.Start.Column - 1] = cell.Text;
                                    }
                                    dataTable.Rows.Add(row);
                                }
                            }
                        }
                        //AutofillColumns(dataTable);
                    }
                    catch (Exception ex) { MessageBox.Show($"Errore: {ex}"); }
                }
            }
            return dataTable;
        }

        public DataTable Query(DateTime startdate, DateTime enddate, string tableName, string connectionString)
        {
            string query = $"SELECT * FROM {tableName}";
            DataTable dt = new DataTable();
            if (!string.IsNullOrEmpty(query) && !string.IsNullOrEmpty(connectionString))
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        try
                        {
                            using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                            {
                                adapter.Fill(dt);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"{ex.Message}");
                        }
                    }
                    connection.Close();
                }
            }
            return dt;
        }
    }
}
