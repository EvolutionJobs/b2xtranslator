using b2xtranslator.DocFileFormat;
using b2xtranslator.OpenXmlLib.WordprocessingML;
using b2xtranslator.OpenXmlLib;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class FooterMapping : DocumentMapping
    {
        private CharacterRange _ftr;

        public FooterMapping(ConversionContext ctx, FooterPart part, CharacterRange ftr)
            : base(ctx, part)
        {
            this._ftr = ftr;
        }
        
        public override void Apply(WordDocument doc)
        {
            this._doc = doc;

            this._writer.WriteStartDocument();
            this._writer.WriteStartElement("w", "ftr", OpenXmlNamespaces.WordprocessingML);

            //convert the footer text
            this._lastValidPapx = this._doc.AllPapxFkps[0].grppapx[0];
            int cp = this._ftr.CharacterPosition;
            int cpMax = this._ftr.CharacterPosition + this._ftr.CharacterCount;

            //the CharacterCount of the footers also counts the guard paragraph mark.
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
