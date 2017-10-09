using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using System.Xml;
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
            Int32 cp = _hdr.CharacterPosition;
            Int32 cpMax = _hdr.CharacterPosition + _hdr.CharacterCount;

            //the CharacterCount of the headers also counts the guard paragraph mark.
            //this additional paragraph mark shall not be converted.
            cpMax--;

            while (cp < cpMax)
            {
                Int32 fc = _doc.PieceTable.FileCharacterPositions[cp];
                ParagraphPropertyExceptions papx = findValidPapx(fc);
                TableInfo tai = new TableInfo(papx);

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
