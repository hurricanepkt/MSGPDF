using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;

namespace ConsoleApplication3
{
    class Program
    {
        static int Main(string[] args)
        {
            var dir = args[0];
            Console.WriteLine("DIR = " + dir);
            var start = (int.Parse(args[1]));
            Console.WriteLine("START = " + start);
            return ConvertXLFiles(dir, start);
        }

        public static int ConvertXLFiles(string outDir, int start)
        {
            DirectoryInfo downloadedMessageInfo = new DirectoryInfo(outDir);
            Workbook workbook = new Workbook();
            foreach (FileInfo file in downloadedMessageInfo.GetFiles("*.xlsx"))
            {
                Console.WriteLine("Loading : " + file.Name); 
                workbook.LoadFromFile(file.FullName, ExcelVersion.Version2013);
                workbook.SaveToFile(string.Format("{0}{1:000}.pdf", outDir, start), FileFormat.PDF);
                file.Delete();
                start++;
            }
            foreach (FileInfo file in downloadedMessageInfo.GetFiles("*.xls"))
            {
                Console.WriteLine("Loading ->" + file.Name);
                workbook.LoadFromFile(file.FullName, ExcelVersion.Version97to2003);
                workbook.SaveToFile(string.Format("{0}{1:000}.pdf", outDir, start), FileFormat.PDF);
                Console.WriteLine("Saving ->" + string.Format("{0}{1:000}.pdf", outDir, start));
                file.Delete();
                start++;
            }
            return start;
        }
    }
}
