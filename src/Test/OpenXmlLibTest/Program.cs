/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System.Text;
using System.IO;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.PresentationML;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;

namespace DIaLOGIKa.b2xtranslator.DocFileFormatTest
{
    class Program
    {
        static void Main(string[] args)
        {
            var doc = WordprocessingDocument.Create(@"C:\tmp\testOpenXmlLib.docx", OpenXmlPackage.DocumentType.Document);

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


            var presentation = PresentationDocument.Create(@"C:\tmp\testOpenXmlLib.pptx", OpenXmlPackage.DocumentType.Document);
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
