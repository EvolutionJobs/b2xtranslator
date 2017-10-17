using b2xtranslator.OpenXmlLib;
using b2xtranslator.DocFileFormat;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class EndnotesMapping : DocumentMapping
    {
        public EndnotesMapping(ConversionContext ctx)
            : base(ctx, ctx.Docx.MainDocumentPart.EndnotesPart)
        {
            this._ctx = ctx;
        }

        public override void Apply(WordDocument doc)
        {
            this._doc = doc;
            int id = 0;

            this._writer.WriteStartElement("w", "endnotes", OpenXmlNamespaces.WordprocessingML);

            int cp = doc.FIB.ccpText + doc.FIB.ccpFtn + doc.FIB.ccpHdr + doc.FIB.ccpAtn;
            int cpEnd = doc.FIB.ccpText + doc.FIB.ccpFtn + doc.FIB.ccpHdr + doc.FIB.ccpAtn + doc.FIB.ccpEdn - 2;
            while (cp < cpEnd)
            {
                this._writer.WriteStartElement("w", "endnote", OpenXmlNamespaces.WordprocessingML);
                this._writer.WriteAttributeString("w", "id", OpenXmlNamespaces.WordprocessingML, id.ToString());
                cp = writeParagraph(cp);
                this._writer.WriteEndElement();
                id++;
            }

            this._writer.WriteEndElement();

            this._writer.Flush();
        }
    }

}
