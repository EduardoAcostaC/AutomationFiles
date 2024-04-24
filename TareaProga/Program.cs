using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using FileOperations;

namespace TareaProga
{
    class Program
    {
        static void Main(string[] args)
        {

            FileHelper fileHelper = new FileHelper();
            string[] baseFileNames = new string[] { "Audit General Engagement Summary YTD Excel_", "Productividad Anual Group Utilization by Submission Date_", "Tiempos cargados Audit SS FY24_" };

            // Rutas de las carpetas
            string[] folderPaths = new string[]
            {
                @"C:\Users\eduar\Downloads\Ingresos",
                @"C:\Users\eduar\Downloads\Productividad",
                @"C:\Users\eduar\Downloads\Tiempos"
            };

            for (int i = 0; i < baseFileNames.Length; i++)
            {
                string baseFileName = baseFileNames[i];
                string folderPath = folderPaths[i];
                string filePattern = baseFileName + "*";
                string[] files = Directory.GetFiles(folderPath, filePattern);

                if (files.Length > 0)
                {
                    // Obtener el archivo más reciente
                    string mostRecentFile = files.OrderByDescending(f => new FileInfo(f).LastWriteTime).FirstOrDefault();

                    // Conectar con Excel
                    Application excelApp = new Application();
                    Workbook excelWorkbook = excelApp.Workbooks.Open(mostRecentFile, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true);
                    Worksheet excelWorksheet = excelWorkbook.Sheets[1];

                    // Conectar con SQL Server
                    string connectionString = @"Data Source = LAPTOP-27K369QS; Initial Catalog = Generacion22; Integrated Security= True;";
                    SqlConnection connection = new SqlConnection(connectionString);
                    connection.Open();

                    // Verificar si la tabla existe y crearla si no
                    string tableName = GetTableName(baseFileName);
                    int colCount = excelWorksheet.UsedRange.Columns.Count;

                    if (!TableExists(connection, tableName))
                    {
                        // Crear tabla en SQL Server
                        string createTableQuery = $"CREATE TABLE {tableName} (LoadDate DATE, ";
                        for (int j = 1; j <= colCount; j++)
                        {
                            string columnName = ((Range)excelWorksheet.Cells[1, j]).Value.ToString();
                            columnName = $"[{columnName}]";
                            createTableQuery += $"{columnName} NVARCHAR(MAX), ";
                        }
                        createTableQuery = createTableQuery.TrimEnd(',', ' ') + ")";
                        SqlCommand createTableCommand = new SqlCommand(createTableQuery, connection);
                        createTableCommand.ExecuteNonQuery();
                    }

                    // Borrar registros del día anterior en la tabla
                    SqlCommand deleteCommand = new SqlCommand($"DELETE FROM {tableName} WHERE LoadDate = CONVERT(DATE, DATEADD(DAY, -1, GETDATE()))", connection);
                    deleteCommand.ExecuteNonQuery();

                    // Leer datos de Excel y guardar en SQL Server
                    int rowCount = excelWorksheet.UsedRange.Rows.Count;
                    for (int j = 2; j <= rowCount; j++)
                    {
                        string insertQuery = $"INSERT INTO {tableName} VALUES ( CONVERT(DATE, GETDATE()), ";
                        for (int k = 1; k <= colCount; k++)
                        {
                            Range cell = (Range)excelWorksheet.Cells[j, k];
                            object value = cell.Value;

                            if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
                            {
                                insertQuery += "NULL, ";
                            }
                            else
                            {
                                insertQuery += $"'{value.ToString()}', ";
                            }
                        }
                        insertQuery = insertQuery.TrimEnd(',', ' ') + ")";
                        SqlCommand insertCommand = new SqlCommand(insertQuery, connection);
                        insertCommand.ExecuteNonQuery();
                    }

                    // Cerrar conexión con Excel y SQL Server
                    excelWorkbook.Close();
                    excelApp.Quit();
                    connection.Close();

                    fileHelper.MoveExcelFile(mostRecentFile, Path.GetFileName(mostRecentFile), folderPath);

                    Console.WriteLine($"Proceso completado para el archivo: {mostRecentFile}");
                }
                else
                {
                    Console.WriteLine($"No se encontraron archivos para {baseFileName}");
                }
            }

        }

        static bool TableExists(SqlConnection connection, string tableName)
        {
            using (SqlCommand command = new SqlCommand($"SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{tableName}'", connection))
            {
                int count = (int)command.ExecuteScalar();
                return count > 0;
            }
        }

        static string GetTableName(string baseFileName)
        {
            if (baseFileName.Contains("General"))
                return "ReporteAuditGeneral";
            else if (baseFileName.Contains("Productividad"))
                return "ReporteAuditProductividad";
            else if (baseFileName.Contains("Tiempos"))
                return "ReporteAuditTiempos";
            else
                throw new ArgumentException("Invalid base file name");
        }
    }
}
