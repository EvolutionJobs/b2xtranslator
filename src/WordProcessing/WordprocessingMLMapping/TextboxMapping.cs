using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class TextboxMapping : DocumentMapping
    {
        public static int TextboxCount = 0;
        private Int32 _textboxIndex;

        public TextboxMapping(ConversionContext ctx, Int32 textboxIndex, ContentPart targetpart, XmlWriter writer)
            : base(ctx, targetpart, writer)
        {
            TextboxCount++;
            _textboxIndex = textboxIndex;
        }

        public TextboxMapping(ConversionContext ctx, ContentPart targetpart, XmlWriter writer)
            : base(ctx, targetpart, writer)
        {
            TextboxCount++;
            _textboxIndex = TextboxCount - 1;
        }

        public override void Apply(WordDocument doc)
        {
            _doc = doc;

            _writer.WriteStartElement("v", "textbox", OpenXmlNamespaces.VectorML);
            _writer.WriteStartElement("w", "txbxContent", OpenXmlNamespaces.WordprocessingML);

            Int32 cp = 0;
            Int32 cpEnd = 0;
            BreakDescriptor bkd = null;
            Int32 txtbxSubdocStart = doc.FIB.ccpText + doc.FIB.ccpFtn + doc.FIB.ccpHdr + doc.FIB.ccpAtn + doc.FIB.ccpEdn;

            if(_targetPart.GetType() == typeof(MainDocumentPart))
            {
                cp = txtbxSubdocStart + doc.TextboxBreakPlex.CharacterPositions[_textboxIndex];
                cpEnd = txtbxSubdocStart + doc.TextboxBreakPlex.CharacterPositions[_textboxIndex + 1];
                bkd = (BreakDescriptor)doc.TextboxBreakPlex.Elements[_textboxIndex];
            }
            if (_targetPart.GetType() == typeof(HeaderPart) || _targetPart.GetType() == typeof(FooterPart))
            {
                txtbxSubdocStart += doc.FIB.ccpTxbx;
                cp = txtbxSubdocStart + doc.TextboxBreakPlexHeader.CharacterPositions[_textboxIndex];
                cpEnd = txtbxSubdocStart + doc.TextboxBreakPlexHeader.CharacterPositions[_textboxIndex + 1];
                bkd = (BreakDescriptor)doc.TextboxBreakPlexHeader.Elements[_textboxIndex];
            }

            //convert the textbox text
            _lastValidPapx = _doc.AllPapxFkps[0].grppapx[0];

            while (cp < cpEnd)
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
            _writer.WriteEndElement();

            _writer.Flush();
        }
    }
}
