using b2xtranslator.DocFileFormat;
using b2xtranslator.OpenXmlLib;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class MainDocumentMapping : DocumentMapping
    {
        public MainDocumentMapping(ConversionContext ctx, ContentPart targetPart)
            : base(ctx, targetPart)
        {
        }

        public override void Apply(WordDocument doc)
        {
            this._doc = doc;

            //start the document
            this._writer.WriteStartDocument();
            this._writer.WriteStartElement("w", "document", OpenXmlNamespaces.WordprocessingML);

            //write namespaces
            this._writer.WriteAttributeString("xmlns", "v", null, OpenXmlNamespaces.VectorML);
            this._writer.WriteAttributeString("xmlns", "o", null, OpenXmlNamespaces.Office);
            this._writer.WriteAttributeString("xmlns", "w10", null, OpenXmlNamespaces.OfficeWord);
            this._writer.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);

            this._writer.WriteStartElement("w", "body", OpenXmlNamespaces.WordprocessingML);

            //convert the document
            this._lastValidPapx = this._doc.AllPapxFkps[0].grppapx[0];
            int cp = 0;
            while (cp < doc.FIB.ccpText)
            {
                int fc = this._doc.PieceTable.FileCharacterPositions[cp];
                var papx = findValidPapx(fc);
                var tai = new TableInfo(papx);

                if (tai.fInTable)
                    //this PAPX is for a table
                    cp = writeTable(cp, tai.iTap);
                else
                    //this PAPX is for a normal paragraph
                    cp = writeParagraph(cp);
            }

            //write the section properties of the body with the last SEPX
            int lastSepxCp = 0;
            foreach (int sepxCp in this._doc.AllSepx.Keys)
                lastSepxCp = sepxCp;
            
            var lastSepx = this._doc.AllSepx[lastSepxCp];
            lastSepx.Convert(new SectionPropertiesMapping(this._writer, this._ctx, this._sectionNr));

            //end the document
            this._writer.WriteEndElement();
            this._writer.WriteEndElement();
            this._writer.WriteEndDocument();

            this._writer.Flush();
        }
    }
}
