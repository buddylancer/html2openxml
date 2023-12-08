using System;
using System.Linq;
using System.IO;
using Ox = DocumentFormat.OpenXml;
using OxP = DocumentFormat.OpenXml.Packaging;
using OxV = DocumentFormat.OpenXml.Validation;
using OxW = DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;

namespace Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            const string filename = "test.docx";
            string html = ResourceHelper.GetString("Resources.CompleteRunTest.html");
            if (File.Exists(filename)) File.Delete(filename);

            using (MemoryStream generatedDocument = new MemoryStream())
            {
                // Uncomment and comment the second using() to open an existing template document
                // instead of creating it from scratch.
                using (var buffer = ResourceHelper.GetStream("Resources.template.docx"))
                {
                    buffer.CopyTo(generatedDocument);
                }

                generatedDocument.Position = 0L;
				using (OxP.WordprocessingDocument package = OxP.WordprocessingDocument.Open(generatedDocument, true))
                //using (WordprocessingDocument package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document))
                {
					OxP.MainDocumentPart mainPart = package.MainDocumentPart;
                    if (mainPart == null)
                    {
                        mainPart = package.AddMainDocumentPart();
						new OxW.Document(new OxW.Body()).Save(mainPart);
                    }

                    HtmlConverter converter = new HtmlConverter(mainPart);
					OxW.Body body = mainPart.Document.Body;

                    converter.ParseHtml(html);
                    mainPart.Document.Save();

                    AssertThatOpenXmlDocumentIsValid(package);
                }

                File.WriteAllBytes(filename, generatedDocument.ToArray());
            }

            System.Diagnostics.Process.Start(filename);
        }

        static void AssertThatOpenXmlDocumentIsValid(OxP.WordprocessingDocument wpDoc)
        {
            var validator = new OxV.OpenXmlValidator(Ox.FileFormatVersions.Office2010);
            var errors = validator.Validate(wpDoc);

            if (!errors.GetEnumerator().MoveNext())
                return;

            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("The document doesn't look 100% compatible with Office 2010.\n");

            Console.ForegroundColor = ConsoleColor.Gray;
            foreach (OxV.ValidationErrorInfo error in errors)
            {
                Console.Write("{0}\n\t{1}", error.Path.XPath, error.Description);
                Console.WriteLine();
            }

            Console.ReadLine();
        }
    }
}