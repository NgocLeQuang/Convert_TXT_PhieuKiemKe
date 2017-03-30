using System;
using System.Collections.Generic;
using System.Text;

namespace ConvertExcelToTXT
{
    public static class Global
    {
        public static  DataDataContext db =new DataDataContext();
        public static string IDProject = "Convert_TXT_PhieuKiemKe";
        public static string Version = "1.0.0";
        public static string UrlUpdateVersion = @"\\10.10.10.254\DE_Viet\2017\Convert_TXT\PhieuKiemKe";
        public static string NameProgram = IDProject + "_V" + Version;
    }
}
