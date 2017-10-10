using b2xtranslator.DocFileFormat;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.OpenXmlLib.WordprocessingML;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class HeaderMapping : DocumentMapping
    {
        private CharacterRange _hdr;

        public HeaderMapping(ConversionContext ctx, HeaderPart part, CharacterRange hdr)
            : base(ctx, part)
        {
            this._hdr = hdr;
        }

        public override void Apply(WordDocument doc)
        {
            this._doc = doc;

            this._writer.WriteStartDocument();
            this._writer.WriteStartElement("w", "hdr", OpenXmlNamespaces.WordprocessingML);

            //convert the header text
            this._lastValidPapx = this._doc.AllPapxFkps[0].grppapx[0];
            int cp = this._hdr.CharacterPosition;
            int cpMax = this._hdr.CharacterPosition + this._hdr.CharacterCount;

            //the CharacterCount of the headers also counts the guard paragraph mark.
            //this additional paragraph mark shall not be converted.
            cpMax--;

            while (cp < cpMax)
            {
                int fc = this._doc.PieceTable.FileCharacterPositions[cp];
                var papx = findValidPapx(fc);
                var tai = new TableInfo(papx);

                if (tai.fInTable)
                {
                    //this PAPX is for a table
                    cp = writeTable(cp, tai.iTap);
                }
                else
                {
                    //this PAPX is for a normal paragraph
                    cp = writeParagraph(cp);
                }
            }

            this._writer.WriteEndElement();
            this._writer.WriteEndDocument();

            this._writer.Flush();

        }
    }
}
