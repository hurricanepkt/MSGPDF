using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClassLibrary1;
using SelectPdf;
using Spire.Doc;
using FileFormat = Spire.Doc.FileFormat;

namespace ConsoleApplication2
{
    class Program
    {
        static void Main(string[] args)
        {
            var outDir = @"C:\Users\Mark\Desktop\trash\MEL\out";
            CleanOutputFolder(outDir);
            ProcessMessage(@"C:\Users\Mark\Desktop\trash\MEL", "Two");
            //Console.ReadKey();
        }

        private static void ProcessMessage(string rootDir, string name)
        {
            var outDir = rootDir + @"\out\" + name + "\\";
            Directory.CreateDirectory(outDir);
            var inDir = rootDir + @"\in\";
            var input = inDir + name + ".msg";
            var d = new MsgReader.Reader();
            var blah = d.ExtractToFolder(input, outDir);
            foreach (var s in blah)
            {
                Console.WriteLine(s);
            }
            CleanOutputFolder(outDir, "*.png");
            CleanOutputFolder(outDir, "*.gif");
            CleanOutputFolder(outDir, "*.jpg");

            var files = ConvertHtmlFiles(outDir, 0);
            files = ConvertDocFiles(outDir, files);
            files = ConvertXLFiles(outDir, files);
        }

        private static int ConvertHtmlFiles(string outDir, int start)
        {
            HtmlToPdf converter = new HtmlToPdf();
            DirectoryInfo downloadedMessageInfo = new DirectoryInfo(outDir);
            foreach (FileInfo file in downloadedMessageInfo.GetFiles("*.htm*"))
            {
                PdfDocument doc = converter.ConvertUrl(file.FullName);
                doc.Save(string.Format("{0}{1:000}.pdf", outDir, start));
                doc.Close();
                file.Delete();
                start++;
            }
            return start;
        }

        private static int ConvertDocFiles(string outDir, int start)
        {
            Document document = new Document();
            DirectoryInfo downloadedMessageInfo = new DirectoryInfo(outDir);
            foreach (FileInfo file in downloadedMessageInfo.GetFiles("*.doc*"))
            {
                
                document.LoadFromFile(file.FullName);
                document.SaveToFile(string.Format("{0}{1:000}.pdf", outDir, start), FileFormat.PDF);
                document.Close();
                file.Delete();
                start++;
            }
            return start;
        }
        private static int ConvertXLFiles(string outDir, int start)
        {
            return Class1.ConvertXLFiles(outDir, start);
        }


        private static void CleanOutputFolder(string outDir, string fileSearch = "")
        {
            DirectoryInfo downloadedMessageInfo = new DirectoryInfo(outDir);

            foreach (FileInfo file in downloadedMessageInfo.GetFiles(fileSearch))
            {
                file.Delete();
            }
            foreach (DirectoryInfo dir in downloadedMessageInfo.GetDirectories())
            {
                try
                {
                    dir.Delete(true);
                }
                catch { }
            }
        }
    }
}
