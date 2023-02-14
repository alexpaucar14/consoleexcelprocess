using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleFileExcelProcessing.Code
{
    internal class Settings
    {
        public bool Get_Rpte_Qualys_Format { get; set; }
        public string RpteQualys_FileName { get; set; }
        public string RpteQualys_FileName_Sheet { get; set; }
        public bool Get_Rpte_Qualys_Diff { get; set; }
        public string ReporteQualys_Diff_FileName1 { get; set; }
        public string ReporteQualys_Diff_FileName1_Sheet { get; set; }
        public string ReporteQualys_Diff_FileName2 { get; set; }
        public string ReporteQualys_Diff_FileName2_Sheet { get; set; }
    }
}
