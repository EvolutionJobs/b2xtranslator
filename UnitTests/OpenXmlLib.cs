using b2xtranslator.OpenXmlLib;
using b2xtranslator.OpenXmlLib.PresentationML;
using b2xtranslator.OpenXmlLib.WordprocessingML;
using NUnit.Framework;
using System.Text;

namespace UnitTests
{
    [TestFixture]
    public class OpenXmlLib
    {
        [Test]
        public void DirectWriteTest()
        {
            var doc = WordprocessingDocument.Create(@"files\testOpenXmlLib.docx", OpenXmlPackage.DocumentType.Document);

            var part = doc.MainDocumentPart;

            const string docXml =
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?> 
<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:body><w:p><w:r><w:t>Hello world!</w:t></w:r></w:p></w:body>
</w:document>";

            var stream = part.GetStream();
            var buf = (new UTF8Encoding()).GetBytes(docXml);
            stream.Write(buf, 0, buf.Length);
        

            doc.Close();


            var presentation = PresentationDocument.Create(@"files\testOpenXmlLib.pptx", OpenXmlPackage.DocumentType.Document);
            var presentationPart = presentation.PresentationPart;

            var slide = presentationPart.AddSlidePart();

            string presentationXml =
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
    <p:presentation xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" 
        xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" 
        xmlns:p=""http://schemas.openxmlformats.org/presentationml/2006/main"" saveSubsetFonts=""1"">
  <p:sldIdLst>
    <p:sldId id=""256"" r:id=""" + slide.RelIdToString + @"""/>
  </p:sldIdLst>
</p:presentation>";
            stream = presentationPart.GetStream();
            buf = (new UTF8Encoding()).GetBytes(presentationXml);
            stream.Write(buf, 0, buf.Length);

            string slideXml =
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<p:sld xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:p=""http://schemas.openxmlformats.org/presentationml/2006/main"">
</p:sld>";

            stream = slide.GetStream();
            buf = (new UTF8Encoding()).GetBytes(slideXml);
            stream.Write(buf, 0, buf.Length);

            presentation.Close();
        }
    }
}
