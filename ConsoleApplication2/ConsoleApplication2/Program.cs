using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using iTextSharp.text.pdf;
using MsgReader;
using SelectPdf;
using Spire.Doc;
using Spire.Pdf;
using FileFormat = Spire.Doc.FileFormat;
using PdfDocument = Spire.Pdf.PdfDocument;

namespace ConsoleApplication2
{
    class Program
    {
        static void Main(string[] args)
        {
            var outDir = @"C:\Users\Mark\Desktop\trash\MEL\out";
            CleanOutputFolder(outDir);
            ProcessMessage(@"C:\Users\Mark\Desktop\trash\MEL", "Three");
            //Console.ReadKey();
        }

        private static void ProcessMessage(string rootDir, string name)
        {
            var outDir = rootDir + @"\out\" + name + "\\";
            Directory.CreateDirectory(outDir);
            var inDir = rootDir + @"\in\";
            var input = inDir + name + ".msg";
            var d = new Reader();
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
            Console.WriteLine("Files Created : " + files);
            Collect(outDir, name);
            Console.WriteLine("Done");
        }

        private static void Collect(string outDir, string name)
        {
           Merger.Collect(outDir, name);
        }

        private static int ConvertHtmlFiles(string outDir, int start)
        {
            HtmlToPdf converter = new HtmlToPdf();
            DirectoryInfo downloadedMessageInfo = new DirectoryInfo(outDir);
            foreach (FileInfo file in downloadedMessageInfo.GetFiles("*.htm*"))
            {
                var doc = converter.ConvertUrl(file.FullName);
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


            // Prepare the process to run
            ProcessStartInfo startProc = new ProcessStartInfo();
            // Enter in the command line arguments, everything you would enter after the executable name itself
            startProc.Arguments = outDir + " " + start;
            // Enter the executable to run, including the complete path
            startProc.FileName = @"XLS\ConsoleApplication3.exe";
            // Do you want to show a console window?
            startProc.WindowStyle = ProcessWindowStyle.Hidden;
            startProc.CreateNoWindow = true;
            int exitCode;


            // Run the external process & wait for it to finish
            using (Process proc = Process.Start(startProc))
            {
                proc.WaitForExit();

                // Retrieve the app's exit code
                exitCode = proc.ExitCode;
            }
            return exitCode;
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
