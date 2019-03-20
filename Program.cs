using System;
using System.IO;
using System.Linq;
using Xceed.Words.NET;

namespace XceedSupport
{
    class Program
    {
        public static Stream ExportToDocX(Stream template)
        {
            Licenser.LicenseKey = Environment.GetEnvironmentVariable("XCEED_WORDS_LICENSE");

            Stream stream = new MemoryStream();
            using (var document = DocX.Load(template))
            {
//                document.InsertParagraph("Inserted Paragraph");
//
//                var consultant = document.Paragraphs.FirstOrDefault(p => p.Text == "<CONSULTANT>");
//                consultant?.Append("Appended Text");
//
//                var candidate = document.Paragraphs.FirstOrDefault(p => p.Text == "<CANDIDATE>");
//                candidate?.Append("Appended Text");

                document.SaveAs(stream);
            }

            return stream;
        }

        static void Main(string[] args)
        {
            var template3 = "templates/Template_Broken.docx";
            var output3 = "output/Output_Broken.docx";

            // This code will produce a corrupt .docx file
            using (var templateStream = File.OpenRead(template3))
            using (var fileStream = File.OpenWrite(output3))
            using (var docxStream = ExportToDocX(templateStream))
            {
                docxStream.Seek(0, SeekOrigin.Begin);
                docxStream.CopyTo(fileStream);
            }

        }
    }
}
