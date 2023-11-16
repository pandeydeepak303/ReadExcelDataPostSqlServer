
using System;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

class Program
{
    static void Main()
    {
        try
        {
            Microsoft.Office.Interop.Excel.Application _excelApp = new Microsoft.Office.Interop.Excel.Application();
            _excelApp.Visible = true;


            string fileName = "G:\\Untitled spreadsheet.xlsx";
            Workbook workbook = _excelApp.Workbooks.Open(fileName,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

            Range excelRange = worksheet.UsedRange;
            object[,] valueArray = (object[,])excelRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);


            string connectionString = "Data Source=DESKTOP-MNV89QG\\SQLEXPRESS;Initial Catalog=ExcelData;Integrated Security=True";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                for (int row = 1; row <= worksheet.UsedRange.Rows.Count; ++row)
                {

                    string empName = valueArray[row, 1]?.ToString();
                    string address = valueArray[row, 2]?.ToString();


                    if (!string.IsNullOrEmpty(empName) && !string.IsNullOrEmpty(address))
                    {
                        string query = "INSERT INTO Employee_Excel (Emp_Name, Address) VALUES (@EmpName, @Address)";

                        using (SqlCommand command = new SqlCommand(query, connection))
                        {
                            command.Parameters.AddWithValue("@EmpName", empName);
                            command.Parameters.AddWithValue("@Address", address);

                            command.ExecuteNonQuery();
                        }
                    }
                }

                Debug.Print("Data inserted into SQL Server successfully");
            }

        }
        catch (Exception ex)
        {
            Debug.Print($"An error occurred: {ex.Message}");
        }
    }
}


