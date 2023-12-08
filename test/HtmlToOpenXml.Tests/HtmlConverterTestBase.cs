using Ox =DocumentFormat.OpenXml;
using OxP =DocumentFormat.OpenXml.Packaging;
using OxW = DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace HtmlToOpenXml.Tests
{
    public abstract class HtmlConverterTestBase
    {
        private System.IO.MemoryStream generatedDocument;
		private OxP.WordprocessingDocument package;

        protected HtmlConverter converter;
		protected OxP.MainDocumentPart mainPart;


        [SetUp]
        public void Init ()
        {
            generatedDocument = new System.IO.MemoryStream();
			package = OxP.WordprocessingDocument.Create(generatedDocument, Ox.WordprocessingDocumentType.Document);

            mainPart = package.MainDocumentPart;
            if (mainPart == null)
            {
                mainPart = package.AddMainDocumentPart();
				new OxW.Document(new OxW.Body()).Save(mainPart);
            }

            this.converter = new HtmlConverter(mainPart);
        }

        [TearDown]
        public void Close ()
        {
            package.Dispose();
            generatedDocument.Dispose();
        }
    }
}