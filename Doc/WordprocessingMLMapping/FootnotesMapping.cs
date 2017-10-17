using b2xtranslator.DocFileFormat;
using b2xtranslator.OpenXmlLib;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class FootnotesMapping : DocumentMapping
    {
        public FootnotesMapping(ConversionContext ctx)
            : base(ctx, ctx.Docx.MainDocumentPart.FootnotesPart)
        {
            this._ctx = ctx;
        }

        public override void Apply(WordDocument doc)
        {
            this._doc = doc;
            int id = 0;

            this._writer.WriteStartElement("w", "footnotes", OpenXmlNamespaces.WordprocessingML);

            int cp = doc.FIB.ccpText;
            while (cp < (doc.FIB.ccpText + doc.FIB.ccpFtn - 2))
            {
                this._writer.WriteStartElement("w", "footnote", OpenXmlNamespaces.WordprocessingML);
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
