using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleFileExcelProcessing.Code
{
    internal class ReporteQualys
    {
        public string IP { get; set; }
        public string HostName { get; set; }
        public string DNS { get; set; }
        public string OS { get; set; }
        public int QID { get; set; }
        public string Title { get; set; }
        public string Vuln_Status { get; set; }
        public string Type { get; set; }
        public int Severity { get; set; }
        public int Port { get; set; }
        public string Protocol { get; set; }
        public string Solution { get; set; }
        public string Results { get; set; }
        public string PCI_Vuln { get; set; }
        public string Category { get; set; }
        public DateTime First_Detected { get; set; }
        public DateTime Last_Detected { get; set; }
        public string Times_Detected { get; set; }
        public string Date_Last_Fixed { get; set; }
        public string First_Reopened { get; set; }
        public string Last_Reopened { get; set; }
        public string Times_Reopened { get; set; }
        public string Herramienta { get; set; }
        public string Ambiente { get; set; }
        public DateTime Fecha { get; set; }
        public string Status { get; set; }
        public string KEY { get; set; }
        public string CODAPP { get; set; }
        public string Squad { get; set; }
        public DateTime? Fecha_Remediacion { get; set; }
        public string Ratificador { get; set; }
        public string Estado_Remediacion { get; set; }
        public string Subject_correo { get; set; }
        public string Ejecutador { get; set; }
        public string Observaciones { get; set; }
        public string Parchado { get; set; }
        public string Obsoleto { get; set; }
        public string Alcance { get; set; }
        public string Responsable { get; set; }
        public string Grupo_Remed { get; set; }
        public Int32 Anio { get; set; }

    }
}
