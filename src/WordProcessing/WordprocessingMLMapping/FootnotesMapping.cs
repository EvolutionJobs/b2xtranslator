using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.Xml;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class FootnotesMapping : DocumentMapping
    {
        public FootnotesMapping(ConversionContext ctx)
            : base(ctx, ctx.Docx.MainDocumentPart.FootnotesPart)
        {
            _ctx = ctx;
        }

        public override void Apply(WordDocument doc)
        {
            _doc = doc;
            int id = 0;

            _writer.WriteStartElement("w", "footnotes", OpenXmlNamespaces.WordprocessingML);

            Int32 cp = doc.FIB.ccpText;
            while (cp < (doc.FIB.ccpText + doc.FIB.ccpFtn - 2))
            {
                _writer.WriteStartElement("w", "footnote", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "id", OpenXmlNamespaces.WordprocessingML, id.ToString());
                cp = writeParagraph(cp);
                _writer.WriteEndElement();
                id++;
            }

            _writer.WriteEndElement();

            _writer.Flush();
        }
    }


}
