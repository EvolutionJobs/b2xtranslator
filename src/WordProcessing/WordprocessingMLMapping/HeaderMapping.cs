using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class HeaderMapping : DocumentMapping
    {
        private CharacterRange _hdr;

        public HeaderMapping(ConversionContext ctx, HeaderPart part, CharacterRange hdr)
            : base(ctx, part)
        {
            _hdr = hdr;
        }

        public override void Apply(WordDocument doc)
        {
            _doc = doc;

            _writer.WriteStartDocument();
            _writer.WriteStartElement("w", "hdr", OpenXmlNamespaces.WordprocessingML);

            //convert the header text
            _lastValidPapx = _doc.AllPapxFkps[0].grppapx[0];
            int cp = _hdr.CharacterPosition;
            int cpMax = _hdr.CharacterPosition + _hdr.CharacterCount;

            //the CharacterCount of the headers also counts the guard paragraph mark.
            //this additional paragraph mark shall not be converted.
            cpMax--;

            while (cp < cpMax)
            {
                int fc = _doc.PieceTable.FileCharacterPositions[cp];
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

            _writer.WriteEndElement();
            _writer.WriteEndDocument();

            _writer.Flush();

        }
    }
}
