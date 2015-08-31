using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2
{
    public class Merger
    {
        public static void Collect(string outDir, string name)
        {
            //DirectoryInfo downloadedMessageInfo = new DirectoryInfo(outDir);
            //var docs = downloadedMessageInfo.GetFiles("*.pdf*").Select(f => new PdfDocument(f.FullName));

            //var m = PdfDocument.MergeFiles(downloadedMessageInfo.GetFiles("*.pdf*").Select(f => f.FullName).ToArray());
            //m.Save(outDir + name + ".pdf", Spire.Pdf.FileFormat.PDF);
            //////open pdf documents    

            ////PdfDocument[] docs = new PdfDocument[files.Length];
            ////for (int i = 0; i < files.Length; i++)
            ////{
            ////    docs[i] = new PdfDocument(files[i]);
            ////}
            //////append document
            ////docs[0].AppendPage(docs[1]);

            //////import PDF pages
            ////for (int i = 0; i < docs[2].Pages.Count; i = i + 2)
            ////{
            ////    docs[0].InsertPage(docs[2], i);
            ////}
        }
    }
}
