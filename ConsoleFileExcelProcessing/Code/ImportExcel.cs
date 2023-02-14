using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net.Http.Headers;
using System.Reflection.Metadata.Ecma335;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Reflection;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using OfficeOpenXml.ConditionalFormatting;

namespace ConsoleFileExcelProcessing.Code
{
    public class ImportExcel
    {
        private static string currentDirectory = Directory.GetCurrentDirectory();
        public void FormateandoReporteQualys(string FileExcel, string sheet)
        {


            string pathFileNameExcel = currentDirectory + @"\Files\" + FileExcel;
            var ReportQualys = ReadReport(pathFileNameExcel, sheet);


            string pathFileNameExcel_Export = currentDirectory + @"\Files\" + FileExcel.Replace(".xlsx", "-Formateado.xlsx");
            ExportExcel(pathFileNameExcel_Export, ReportQualys.Distinct().ToList());
        }

        public void ProcessReportQualys(string fileNameExcel, string Sheet, string fileNameExcel2, string Sheet2)
        {
            string pathFileNameExcel = currentDirectory + @"\Files\" + fileNameExcel;
            var ReportPrincipal = ReadReport(pathFileNameExcel, Sheet);

            var lstDictionary = ReadDictionary(pathFileNameExcel, "Diccionario");

            string pathFileNameExcel2 = currentDirectory + @"\Files\" + fileNameExcel2;
            var NewReportQualys = ReadReport(pathFileNameExcel2, Sheet2);

            List<ReporteQualys> lstReportQualysComplete = new List<ReporteQualys>();

            //---
            foreach (var item in NewReportQualys)
            {
                var obj = ReportPrincipal.FirstOrDefault(r => r.QID.Equals(item.QID) && r.IP.Equals(item.IP));
                if (obj != null)
                {
                    item.Status = item.Vuln_Status;
                    item.KEY = obj.KEY;

                    item.Squad = obj.Squad;
                    item.Ambiente = obj.Ambiente;
                    item.Herramienta = obj.Herramienta;

                    //if (item.Squad.Equals("Otros", StringComparison.CurrentCultureIgnoreCase))
                    //{
                        item.Squad = lstDictionary.FirstOrDefault(f => f.IP == item.IP) != null ? lstDictionary.FirstOrDefault(f => f.IP == item.IP).Squad : "Otros";
                    //}
                    //if (item.Ambiente.Equals("Otros", StringComparison.CurrentCultureIgnoreCase) ||
                    //    item.Ambiente.Equals("-"))
                    //{
                        item.Ambiente = lstDictionary.FirstOrDefault(f => f.IP == item.IP) != null ? lstDictionary.FirstOrDefault(f => f.IP == item.IP).Enviorement : "Otros";
                    //}
                    //if (item.Herramienta.Equals("Otros", StringComparison.CurrentCultureIgnoreCase))
                    //{
                        item.Herramienta = lstDictionary.FirstOrDefault(f => f.IP == item.IP) != null ? lstDictionary.FirstOrDefault(f => f.IP == item.IP).Tool : "Otros";
                    //}
                    //if (string.IsNullOrEmpty(item.HostName))
                    //{
                        item.HostName = lstDictionary.FirstOrDefault(f => f.IP == item.IP) != null ? lstDictionary.FirstOrDefault(f => f.IP == item.IP).HostName : item.IP;
                    //}
                    item.Fecha = obj.Fecha;
                    //Recurrentes
                    item.Fecha_Remediacion = obj.Fecha_Remediacion;
                    item.Ratificador = obj.Ratificador;
                    item.Estado_Remediacion = obj.Estado_Remediacion;
                    item.Subject_correo = obj.Subject_correo;
                    item.Ejecutador = obj.Ejecutador;
                    item.Observaciones = obj.Observaciones;
                    item.Parchado = obj.Parchado;
                    item.Anio = obj.First_Detected.Year;

                    if (!string.IsNullOrEmpty(obj.Obsoleto))
                    {
                        item.Obsoleto = obj.Obsoleto;
                        item.Alcance = obj.Alcance;
                        item.Responsable = obj.Responsable;
                        item.Grupo_Remed = obj.Grupo_Remed;
                    }

                    lstReportQualysComplete.Add(item);
                }
                else
                {
                    item.Squad = lstDictionary.FirstOrDefault(f => f.IP == item.IP) != null ? lstDictionary.FirstOrDefault(f => f.IP == item.IP).Squad : "Otros";
                    item.Ambiente = lstDictionary.FirstOrDefault(f => f.IP == item.IP) != null ? lstDictionary.FirstOrDefault(f => f.IP == item.IP).Enviorement : "Otros";
                    item.Herramienta = lstDictionary.FirstOrDefault(f => f.IP == item.IP) != null ? lstDictionary.FirstOrDefault(f => f.IP == item.IP).Tool : "Otros";
                    item.HostName = lstDictionary.FirstOrDefault(f => f.IP == item.IP) != null ? lstDictionary.FirstOrDefault(f => f.IP == item.IP).HostName : item.IP;

                    item.KEY = String.Concat(item.IP, item.QID);
                    item.Status = item.Vuln_Status;
                    item.Fecha = DateTime.Now;
                    item.Anio = item.First_Detected.Year;
                    //item.Observaciones = String.Concat(item.Observaciones, " [ Cerrado el ", DateTime.Now, " ]");
                    lstReportQualysComplete.Add((ReporteQualys)item);
                }
            }



            foreach (var item in ReportPrincipal)
            {
                if (!NewReportQualys.Any(r => r.QID.Equals(item.QID) && r.IP.Equals(item.IP)))
                {
                    if (!item.Status.Equals("Closed"))
                    {
                        item.Status = "Closed";
                        item.Fecha = DateTime.Now;
                        item.Observaciones = String.Concat(item.Observaciones, " [ Cerrado el ", DateTime.Now, " ]");
                        if (string.IsNullOrEmpty(item.HostName))
                        {
                            item.HostName = lstDictionary.FirstOrDefault(f => f.IP == item.IP) != null ? lstDictionary.FirstOrDefault(f => f.IP == item.IP).HostName : item.IP;
                        }
                    }
                    lstReportQualysComplete.Add((ReporteQualys)item);


                    //--------

                }
            }

            string pathFileNameExcelFinal = currentDirectory + @"\Files\" + fileNameExcel.Replace(".xlsx", "-BD.xlsx");
            //ExportExcel(pathFileNameExcelFinal, lstReportQualysComplete.Distinct().ToList());
            ExportExcel(pathFileNameExcelFinal, lstReportQualysComplete);
        }

        private List<ReporteQualys> ReadReport(string PathFilenameExcel, string sheet)
        {
            List<ReporteQualys> ReporteQualysList = null;
            try
            {
                System.Data.OleDb.OleDbConnection myConnection = new System.Data.OleDb.OleDbConnection(
                                                                 "Provider=Microsoft.ACE.OLEDB.12.0; " +
                                                                 "data source='" + PathFilenameExcel + "';" +
                                                                 "Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\" ");
                myConnection.Open();

                var sheetName = sheet + "$";
                DataTable dt = makeDataTableFromSheetName(PathFilenameExcel, sheetName);

                ReporteQualysList = new List<ReporteQualys>();

                if (PathFilenameExcel.Contains("Reporte VUL Qualys", StringComparison.InvariantCultureIgnoreCase))
                {
                    ReporteQualysList = (from DataRow dr in dt.Rows
                                         select new ReporteQualys()
                                         {
                                             IP = dr["IP"] != DBNull.Value ? dr["IP"].ToString() : string.Empty,
                                             HostName = dr["HostName"] != DBNull.Value ? dr["HostName"].ToString() : string.Empty,
                                             DNS = dr["DNS"] != DBNull.Value ? dr["DNS"].ToString() : string.Empty,
                                             OS = dr["OS"] != DBNull.Value ? dr["OS"].ToString() : string.Empty,
                                             QID = dr["QID"] != DBNull.Value ? Convert.ToInt32(dr["QID"]) : 0,
                                             Title = dr["Title"] != DBNull.Value ? dr["Title"].ToString() : string.Empty,
                                             Vuln_Status = dr["Vuln Status"] != DBNull.Value ? dr["Vuln Status"].ToString() : string.Empty,
                                             Type = dr["Type"] != DBNull.Value ? dr["Type"].ToString() : string.Empty,
                                             Severity = dr["Severity"] != DBNull.Value ? Convert.ToInt32(dr["Severity"]) : 0,
                                             Port = dr["Port"] != DBNull.Value ? Convert.ToInt32(dr["Port"]) : 0,
                                             Protocol = dr["Protocol"] != DBNull.Value ? dr["Protocol"].ToString() : string.Empty,
                                             Solution = dr["Solution"] != DBNull.Value ? dr["Solution"].ToString() : string.Empty,
                                             Results = dr["Results"] != DBNull.Value ? dr["Results"].ToString() : string.Empty,
                                             PCI_Vuln = dr["PCI Vuln"] != DBNull.Value ? dr["PCI Vuln"].ToString() : string.Empty,
                                             Category = dr["Category"] != DBNull.Value ? dr["Category"].ToString() : string.Empty,
                                             First_Detected = dr["First Detected"] != DBNull.Value ? Convert.ToDateTime(dr["First Detected"]) : DateTime.Now,
                                             Last_Detected = dr["Last Detected"] != DBNull.Value ? Convert.ToDateTime(dr["Last Detected"]) : DateTime.Now,
                                             Times_Detected = dr["Times Detected"] != DBNull.Value ? dr["Times Detected"].ToString() : string.Empty,
                                             Date_Last_Fixed = dr["Date Last Fixed"] != DBNull.Value ? dr["Date Last Fixed"].ToString() : string.Empty,
                                             First_Reopened = dr["First Reopened"] != DBNull.Value ? dr["First Reopened"].ToString() : string.Empty,
                                             Last_Reopened = dr["Last Reopened"] != DBNull.Value ? dr["Last Reopened"].ToString() : string.Empty,
                                             Times_Reopened = dr["Times Reopened"] != DBNull.Value ? dr["Times Reopened"].ToString() : string.Empty,
                                             Herramienta = dr["Herramienta"] != DBNull.Value ? dr["Herramienta"].ToString() : string.Empty,
                                             Ambiente = dr["Ambiente"] != DBNull.Value ? dr["Ambiente"].ToString() : string.Empty,
                                             Fecha = dr["Fecha"] != DBNull.Value ? Convert.ToDateTime(dr["Fecha"]) : DateTime.Now,
                                             Status = dr["Status"] != DBNull.Value ? dr["Status"].ToString() : string.Empty,
                                             KEY = dr["KEY"] != DBNull.Value ? dr["KEY"].ToString() : string.Empty,
                                             CODAPP = dr["CODAPP"] != DBNull.Value ? dr["CODAPP"].ToString() : string.Empty,
                                             Squad = dr["Squad"] != DBNull.Value ? dr["Squad"].ToString() : string.Empty,
                                             Fecha_Remediacion = dr["Fecha Remediacion"] != DBNull.Value ? Convert.ToDateTime(dr["Fecha Remediacion"]) : null,
                                             Ratificador = dr["Ratificador"] != DBNull.Value ? dr["Ratificador"].ToString() : string.Empty,
                                             Estado_Remediacion = dr["Estado Remediacion"] != DBNull.Value ? dr["Estado Remediacion"].ToString() : string.Empty,
                                             Subject_correo = dr["Subject correo"] != DBNull.Value ? dr["Subject correo"].ToString() : string.Empty,
                                             Ejecutador = dr["Ejecutador"] != DBNull.Value ? dr["Ejecutador"].ToString() : string.Empty,
                                             Observaciones = dr["Observaciones"] != DBNull.Value ? dr["Observaciones"].ToString() : string.Empty,
                                             Parchado = dr["Parchado"] != DBNull.Value ? dr["Parchado"].ToString() : string.Empty,
                                             
                                             Obsoleto = dr["Obsoleto"] != DBNull.Value ? dr["Obsoleto"].ToString() : string.Empty,
                                             Alcance = dr["Alcance"] != DBNull.Value ? dr["Alcance"].ToString() : string.Empty,
                                             Responsable = dr["Responsable"] != DBNull.Value ? dr["Responsable"].ToString() : string.Empty,
                                             Grupo_Remed = dr["Grupo Remed"] != DBNull.Value ? dr["Grupo Remed"].ToString() : string.Empty,
                                             Anio = dr["First Detected"] != DBNull.Value ? Convert.ToDateTime(dr["First Detected"]).Year : DateTime.Now.Year,
                                         }).ToList();
                }
                else
                {
                    ReporteQualysList = (from DataRow dr in dt.Rows
                                         select new ReporteQualys()
                                         {
                                             IP = dr["IP"] != DBNull.Value ? dr["IP"].ToString() : string.Empty,
                                             DNS = dr["DNS"] != DBNull.Value ? dr["DNS"].ToString() : string.Empty,
                                             OS = dr["OS"] != DBNull.Value ? dr["OS"].ToString() : string.Empty,
                                             QID = dr["QID"] != DBNull.Value ? Convert.ToInt32(dr["QID"]) : 0,
                                             Title = dr["Title"] != DBNull.Value ? dr["Title"].ToString() : string.Empty,
                                             Vuln_Status = dr["Vuln Status"] != DBNull.Value ? dr["Vuln Status"].ToString() : string.Empty,
                                             Type = dr["Type"] != DBNull.Value ? dr["Type"].ToString() : string.Empty,
                                             Severity = dr["Severity"] != DBNull.Value ? Convert.ToInt32(dr["Severity"]) : 0,
                                             Port = dr["Port"] != DBNull.Value ? Convert.ToInt32(dr["Port"]) : 0,
                                             Protocol = dr["Protocol"] != DBNull.Value ? dr["Protocol"].ToString() : string.Empty,
                                             Solution = dr["Solution"] != DBNull.Value ? dr["Solution"].ToString() : string.Empty,
                                             Results = dr["Results"] != DBNull.Value ? dr["Results"].ToString() : string.Empty,
                                             PCI_Vuln = dr["PCI Vuln"] != DBNull.Value ? dr["PCI Vuln"].ToString() : string.Empty,
                                             Category = dr["Category"] != DBNull.Value ? dr["Category"].ToString() : string.Empty,
                                             First_Detected = dr["First Detected"] != DBNull.Value ? Convert.ToDateTime(dr["First Detected"]) : DateTime.Now,
                                             Last_Detected = dr["Last Detected"] != DBNull.Value ? Convert.ToDateTime(dr["Last Detected"]) : DateTime.Now,

                                             //First_Detected = dr["First Detected"] != DBNull.Value ? DateTime.ParseExact(dr["First Detected"].ToString(),"MM/dd/yyyy",null) : DateTime.Now,
                                             //Last_Detected = dr["Last Detected"] != DBNull.Value ? DateTime.ParseExact(dr["Last Detected"].ToString(), "MM/dd/yyyy", null) : DateTime.Now,

                                             Times_Detected = dr["Times Detected"] != DBNull.Value ? dr["Times Detected"].ToString() : string.Empty,
                                             Date_Last_Fixed = dr["Date Last Fixed"] != DBNull.Value ? dr["Date Last Fixed"].ToString() : string.Empty,
                                             First_Reopened = dr["First Reopened"] != DBNull.Value ? dr["First Reopened"].ToString() : string.Empty,
                                             Last_Reopened = dr["Last Reopened"] != DBNull.Value ? dr["Last Reopened"].ToString() : string.Empty,
                                             Times_Reopened = dr["Times Reopened"] != DBNull.Value ? dr["Times Reopened"].ToString() : string.Empty,
                                             CODAPP = dr["CODAPP"] != DBNull.Value ? dr["CODAPP"].ToString() : string.Empty

                                             /*Obsoleto = dr["Obsoleto"] != DBNull.Value ? dr["Obsoleto"].ToString() : string.Empty,
                                             Alcance = dr["Alcance"] != DBNull.Value ? dr["Alcance"].ToString() : string.Empty,
                                             Responsable = dr["Responsable"] != DBNull.Value ? dr["Responsable"].ToString() : string.Empty,
                                             Grupo_Remed = dr["Grupo Remed"] != DBNull.Value ? dr["Grupo Remed"].ToString() : string.Empty*/
                                         }).ToList();
                }



                return ReporteQualysList;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                return ReporteQualysList;
            }

        }

        private List<DictionaryTools> ReadDictionary(string PathFilenameExcel, string sheet)
        {
            List<DictionaryTools> DictionaryToolsList = null;
            try
            {
                System.Data.OleDb.OleDbConnection myConnection = new System.Data.OleDb.OleDbConnection(
                                                                 "Provider=Microsoft.ACE.OLEDB.12.0; " +
                                                                 "data source='" + PathFilenameExcel + "';" +
                                                                 "Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\" ");
                myConnection.Open();

                var sheetName = sheet + "$";
                DataTable dt = makeDataTableFromSheetName(PathFilenameExcel, sheetName);

                DictionaryToolsList = new List<DictionaryTools>();

                DictionaryToolsList = (from DataRow dr in dt.Rows
                                       select new DictionaryTools()
                                       {
                                           IP = dr["IP"] != DBNull.Value ? dr["IP"].ToString() : string.Empty,
                                           Tool = dr["Tool"] != DBNull.Value ? dr["Tool"].ToString() : string.Empty,
                                           Details = dr["Details"] != DBNull.Value ? dr["Details"].ToString() : string.Empty,
                                           Suscription = dr["Suscription"] != DBNull.Value ? dr["Suscription"].ToString() : string.Empty,
                                           Enviorement = dr["Enviorement"] != DBNull.Value ? dr["Enviorement"].ToString() : string.Empty,
                                           HostName = dr["HostName"] != DBNull.Value ? dr["HostName"].ToString() : string.Empty,
                                           Squad = dr["Squad"] != DBNull.Value ? dr["Squad"].ToString() : string.Empty
                                       }).ToList();



                return DictionaryToolsList;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                return DictionaryToolsList;
            }

        }

        private void ExportExcel(string filePathExcelExport, List<ReporteQualys> lstReporteQualys)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                int contWorksheet = 0;
                int initalCol = 1;
                int initialRow = 0;

                excelPackage.Workbook.Properties.Author = "BCP";
                excelPackage.Workbook.Properties.Title = "Qualys";

                // Creamos un Worksheet
                excelPackage.Workbook.Worksheets.Add("Reporte Qualys");

                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[contWorksheet];
                ws.Cells.Style.Font.Size = 10;
                ws.Cells.Style.Font.Name = "Arial";

                int colIndex = initalCol;
                int rowIndex = initialRow;

                PropertyInfo[] properties = typeof(ReporteQualys).GetProperties();
                foreach (var property in properties)
                {
                    string NombreAtributo = property.Name;
                    ws.Cells[rowIndex + 1, colIndex].Value = NombreAtributo;
                    ws.Cells[rowIndex + 1, colIndex].Style.Font.Bold = true;
                    ws.Cells[rowIndex + 1, colIndex].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ws.Cells[rowIndex + 1, colIndex].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells[rowIndex + 1, colIndex].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    ws.Cells[rowIndex + 1, colIndex].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[rowIndex + 1, colIndex].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ws.Cells[rowIndex + 1, colIndex].Style.Font.Color.SetColor(Color.White);
                    ws.Cells[rowIndex + 1, colIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[rowIndex + 1, colIndex].Style.Fill.BackgroundColor.SetColor(Color.Green);
                    colIndex++;
                }

                rowIndex++;
                colIndex = initalCol;

                foreach (var objvul in lstReporteQualys)
                {
                    colIndex = initalCol;
                    rowIndex++;

                    foreach (var property in properties)
                    {
                        switch (property.Name)
                        {
                            case "First_Detected":
                            case "Last_Detected":
                            case "Fecha":
                            case "Fecha_Remediacion":
                                if (property.GetValue(objvul) != null)
                                {
                                    ws.Cells[rowIndex, colIndex].Value = (DateTime)property.GetValue(objvul);
                                }
                                ws.Cells[rowIndex, colIndex].Style.Numberformat.Format = "dd/MM/yyyy hh:mm";
                                break;
                            case "QID":
                            case "Port":
                            case "Severity":
                            case "Anio":
                                ws.Cells[rowIndex, colIndex].Value = (int)property.GetValue(objvul);
                                break;
                            default:
                                ws.Cells[rowIndex, colIndex].Value = (string)property.GetValue(objvul);
                                break;
                        }


                        ws.Cells[rowIndex, colIndex].Style.Font.Size = 11;
                        ws.Cells[rowIndex, colIndex].Style.Font.Name = "Calibri";
                        ws.Cells[rowIndex, colIndex].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        colIndex++;
                    }

                }
                Byte[] bin = excelPackage.GetAsByteArray();


                String file = filePathExcelExport;
                File.WriteAllBytes(file, bin);
            }
        }
        public void ProcessingExcel()
        {
            string filename = @"D:\Code\demoExcel.xlsx";
            System.Data.OleDb.OleDbConnection myConnection = new System.Data.OleDb.OleDbConnection(
                        "Provider=Microsoft.ACE.OLEDB.12.0; " +
                         "data source='" + filename + "';" +
                            "Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\" ");
            myConnection.Open();
            DataTable mySheets = myConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            //DataSet ds = new DataSet();
            DataTable dt;

            List<ReporteQualys> ReporteQualysList = new List<ReporteQualys>();
            var sheetNames = mySheets.Rows[0]["TABLE_NAME"].ToString();
            dt = makeDataTableFromSheetName(filename, sheetNames);


            ReporteQualysList = (from DataRow dr in dt.Rows
                                 select new ReporteQualys()
                                 {
                                     IP = dr["IP"].ToString(),
                                     DNS = dr["DNS"].ToString(),
                                     OS = dr["OS"].ToString(),
                                     QID = Convert.ToInt32(dr["QID"]),
                                     Title = dr["Title"].ToString(),
                                     Vuln_Status = dr["Vuln Status"].ToString(),
                                     Type = dr["Type"].ToString(),
                                     Severity = Convert.ToInt32(dr["Severity"]),
                                     Port = Convert.ToInt32(dr["Port"]),
                                     Protocol = dr["Protocol"].ToString(),
                                     Solution = dr["Solution"].ToString(),
                                     Results = dr["Results"].ToString(),
                                     PCI_Vuln = dr["PCI Vuln"].ToString(),
                                     Category = dr["Category"].ToString(),
                                     First_Detected = Convert.ToDateTime(dr["First Detected"]),
                                     Last_Detected = Convert.ToDateTime(dr["Last Detected"]),
                                     Times_Detected = dr["Times Detected"].ToString(),
                                     Date_Last_Fixed = dr["Date Last Fixed"].ToString(),
                                     First_Reopened = dr["First Reopened"].ToString(),
                                     Last_Reopened = dr["Last Reopened"].ToString(),
                                     Times_Reopened = dr["Times Reopened"].ToString(),
                                     Herramienta = dr["Herramienta"].ToString(),
                                     Ambiente = dr["Ambiente"].ToString(),
                                     Fecha = Convert.ToDateTime(dr["Fecha"]),
                                     Status = dr["Status"].ToString(),
                                     KEY = dr["KEY"].ToString(),
                                     Squad = dr["Squad"].ToString()
                                 }).ToList();



            //---
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                int contWorksheet = 0;
                int initalCol = 1;
                int initialRow = 0;

                excelPackage.Workbook.Properties.Author = "BCP";
                excelPackage.Workbook.Properties.Title = "Qualys";

                // Creamos un Worksheet
                excelPackage.Workbook.Worksheets.Add("Reporte Qualys");

                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[contWorksheet];
                ws.Cells.Style.Font.Size = 10;
                ws.Cells.Style.Font.Name = "Arial";

                int colIndex = initalCol;
                int rowIndex = initialRow;

                PropertyInfo[] properties = typeof(ReporteQualys).GetProperties();
                foreach (var property in properties)
                {
                    string NombreAtributo = property.Name;
                    ws.Cells[rowIndex + 1, colIndex].Value = NombreAtributo;
                    ws.Cells[rowIndex + 1, colIndex].Style.Font.Bold = true;
                    ws.Cells[rowIndex + 1, colIndex].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ws.Cells[rowIndex + 1, colIndex].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells[rowIndex + 1, colIndex].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    ws.Cells[rowIndex + 1, colIndex].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[rowIndex + 1, colIndex].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ws.Cells[rowIndex + 1, colIndex].Style.Font.Color.SetColor(Color.White);
                    ws.Cells[rowIndex + 1, colIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[rowIndex + 1, colIndex].Style.Fill.BackgroundColor.SetColor(Color.Green);
                    colIndex++;
                }

                rowIndex++;
                colIndex = initalCol;

                foreach (var objvul in ReporteQualysList)
                {
                    colIndex = initalCol;
                    rowIndex++;

                    foreach (var property in properties)
                    {
                        switch (property.Name)
                        {
                            case "First_Detected":
                            case "Last_Detected":
                            case "Fecha":
                                ws.Cells[rowIndex, colIndex].Value = (DateTime)property.GetValue(objvul);
                                ws.Cells[rowIndex, colIndex].Style.Numberformat.Format = "dd/MM/yyyy hh:mm";
                                break;
                            case "QID":
                            case "Port":
                            case "Severity":
                                ws.Cells[rowIndex, colIndex].Value = (int)property.GetValue(objvul);
                                break;
                            default:
                                ws.Cells[rowIndex, colIndex].Value = (string)property.GetValue(objvul);
                                break;
                        }


                        ws.Cells[rowIndex, colIndex].Style.Font.Size = 11;
                        ws.Cells[rowIndex, colIndex].Style.Font.Name = "Calibri";
                        ws.Cells[rowIndex, colIndex].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        colIndex++;
                    }

                }
                Byte[] bin = excelPackage.GetAsByteArray();

                string filePathExport = @"D:\Code\demoExcelExport.xlsx";
                String file = filePathExport;
                File.WriteAllBytes(file, bin);
            }
            //---

            var ssss = "";
            //}



        }


        public DataTable makeDataTableFromSheetName(string filename, string sheetName)
        {
            System.Data.OleDb.OleDbConnection myConnection = new System.Data.OleDb.OleDbConnection(
            "Provider=Microsoft.ACE.OLEDB.12.0; " +
            "data source='" + filename + "';" +
            "Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\" ");

            DataTable dtImport = new DataTable();
            System.Data.OleDb.OleDbDataAdapter myImportCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [" + sheetName + "]", myConnection);
            myImportCommand.Fill(dtImport);
            return dtImport;
        }


        public void ProcessReportQualys_Backup(string fileNameExcel, string Sheet, string fileNameExcel2, string Sheet2)
        {
            string pathFileNameExcel = currentDirectory + @"\Files\" + fileNameExcel;
            var ReportPrincipal = ReadReport(pathFileNameExcel, Sheet);

            var lstDictionary = ReadDictionary(pathFileNameExcel, "Diccionario");

            string pathFileNameExcel2 = currentDirectory + @"\Files\" + fileNameExcel2;
            var ReportQualys = ReadReport(pathFileNameExcel2, Sheet2);

            List<ReporteQualys> lstReportQualysComplete = new List<ReporteQualys>();

            foreach (var item in ReportPrincipal)
            {
                var obj = ReportQualys.FirstOrDefault(r => r.QID.Equals(item.QID) && r.IP.Equals(item.IP));
                if (obj != null)
                {
                    obj.Status = obj.Vuln_Status;
                    obj.KEY = item.KEY;

                    obj.Squad = item.Squad;
                    obj.Ambiente = item.Ambiente;
                    obj.Herramienta = item.Herramienta;

                    if (obj.Squad.Equals("Otros", StringComparison.CurrentCultureIgnoreCase))
                    {
                        obj.Squad = lstDictionary.FirstOrDefault(f => f.IP == item.IP) != null ? lstDictionary.FirstOrDefault(f => f.IP == item.IP).Squad : "Otros";
                    }
                    if (obj.Ambiente.Equals("Otros", StringComparison.CurrentCultureIgnoreCase) ||
                        obj.Ambiente.Equals("-"))
                    {
                        obj.Ambiente = lstDictionary.FirstOrDefault(f => f.IP == item.IP) != null ? lstDictionary.FirstOrDefault(f => f.IP == item.IP).Enviorement : "Otros";
                    }
                    if (obj.Herramienta.Equals("Otros", StringComparison.CurrentCultureIgnoreCase))
                    {
                        obj.Herramienta = lstDictionary.FirstOrDefault(f => f.IP == item.IP) != null ? lstDictionary.FirstOrDefault(f => f.IP == item.IP).Tool : "Otros";
                    }
                    if (string.IsNullOrEmpty(obj.HostName))
                    {
                        obj.HostName = lstDictionary.FirstOrDefault(f => f.IP == item.IP) != null ? lstDictionary.FirstOrDefault(f => f.IP == item.IP).HostName : item.IP;
                    }

                    obj.Fecha = item.Fecha;
                    //Recurrentes
                    obj.Fecha_Remediacion = item.Fecha_Remediacion;
                    obj.Ratificador = item.Ratificador;
                    obj.Estado_Remediacion = item.Estado_Remediacion;
                    obj.Subject_correo = item.Subject_correo;
                    obj.Ejecutador = item.Ejecutador;
                    obj.Observaciones = item.Observaciones;
                    lstReportQualysComplete.Add(obj);
                }
                else
                {
                    item.Status = "Closed";
                    item.Fecha = DateTime.Now;
                    item.Observaciones = String.Concat(item.Observaciones, " [ Cerrado el ", DateTime.Now, " ]");
                    if (string.IsNullOrEmpty(item.HostName))
                    {
                        item.HostName = lstDictionary.FirstOrDefault(f => f.IP == item.IP) != null ? lstDictionary.FirstOrDefault(f => f.IP == item.IP).HostName : item.IP;
                    }
                    lstReportQualysComplete.Add((ReporteQualys)item);
                }
            }

            foreach (var item in ReportQualys)
            {
                if (!ReportPrincipal.Any(r => r.QID.Equals(item.QID) && r.IP.Equals(item.IP)))
                {
                    //item.Status = item.Vuln_Status;
                    item.Status = "New";
                    item.Fecha = DateTime.Now;
                    item.KEY = String.Concat(item.IP, item.QID);
                    item.Squad = lstDictionary.FirstOrDefault(f => f.IP == item.IP) != null ? lstDictionary.FirstOrDefault(f => f.IP == item.IP).Squad : "Otros";
                    item.Ambiente = lstDictionary.FirstOrDefault(f => f.IP == item.IP) != null ? lstDictionary.FirstOrDefault(f => f.IP == item.IP).Enviorement : "Otros";
                    item.Herramienta = lstDictionary.FirstOrDefault(f => f.IP == item.IP) != null ? lstDictionary.FirstOrDefault(f => f.IP == item.IP).Tool : "Otros";
                    item.HostName = lstDictionary.FirstOrDefault(f => f.IP == item.IP) != null ? lstDictionary.FirstOrDefault(f => f.IP == item.IP).HostName : item.IP;
                    item.Observaciones = "New VUL";
                    lstReportQualysComplete.Add((ReporteQualys)item);
                }
            }

            string pathFileNameExcelFinal = currentDirectory + @"\Files\" + fileNameExcel.Replace(".xlsx", "-BD.xlsx");
            //ExportExcel(pathFileNameExcelFinal, lstReportQualysComplete.Distinct().ToList());
            ExportExcel(pathFileNameExcelFinal, lstReportQualysComplete);
        }
        #region "Functions"
        private static List<T> ConvertDataTable<T>(DataTable dt)
        {
            List<T> data = new List<T>();
            foreach (DataRow row in dt.Rows)
            {
                T item = GetItem<T>(row);
                data.Add(item);
            }
            return data;
        }
        private static T GetItem<T>(DataRow dr)
        {
            Type temp = typeof(T);
            T obj = Activator.CreateInstance<T>();

            foreach (DataColumn column in dr.Table.Columns)
            {
                foreach (PropertyInfo pro in temp.GetProperties())
                {
                    if (pro.Name == column.ColumnName)
                        pro.SetValue(obj, dr[column.ColumnName], null);
                    else
                        continue;
                }
            }
            return obj;
        }


        #endregion
    }
}
