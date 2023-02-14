// See https://aka.ms/new-console-template for more information
using ConsoleFileExcelProcessing.Code;
using Microsoft.Extensions.Configuration;
using System.Data;

Console.WriteLine("Procesando...");


IConfiguration config = new ConfigurationBuilder()
    .AddJsonFile("appsettings.json")
    .AddEnvironmentVariables()
    .Build();
Settings settings = config.GetRequiredSection("Settings").Get<Settings>();

if (settings.Get_Rpte_Qualys_Format)
{
    Console.WriteLine("Formateando Reporte Qualys");
    new ImportExcel().FormateandoReporteQualys(settings.RpteQualys_FileName, settings.RpteQualys_FileName_Sheet);
}

if (settings.Get_Rpte_Qualys_Diff)
{
    Console.WriteLine("Generando Reporte DB");
    new ImportExcel().ProcessReportQualys(settings.ReporteQualys_Diff_FileName1, settings.ReporteQualys_Diff_FileName1_Sheet, settings.ReporteQualys_Diff_FileName2, settings.ReporteQualys_Diff_FileName2_Sheet);
}


Console.WriteLine("Completado!");

