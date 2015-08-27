using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;

namespace ClassLibrary1
{
    public class Class1
    {
        public static int ConvertXLFiles(string outDir, int start)
        {
            DirectoryInfo downloadedMessageInfo = new DirectoryInfo(outDir);
            Workbook workbook = new Workbook();
            foreach (FileInfo file in downloadedMessageInfo.GetFiles("*.xlsx"))
            {
                workbook.LoadFromFile(file.FullName, ExcelVersion.Version2013);
                workbook.SaveToFile(string.Format("{0}{1:000}.pdf", outDir, start), FileFormat.PDF);
                //file.Delete();
                start++;
            }
            foreach (FileInfo file in downloadedMessageInfo.GetFiles("*.xls"))
            {
                workbook.LoadFromFile(file.FullName, ExcelVersion.Version97to2003);
                workbook.SaveToFile(string.Format("{0}{1:000}.pdf", outDir, start), FileFormat.PDF);
                //file.Delete();
                start++;
            }
            return start;
        }
    }


}
