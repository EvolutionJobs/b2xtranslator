using System.Xml;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.DocFileFormat;
using b2xtranslator.OpenXmlLib.WordprocessingML;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class TextboxMapping : DocumentMapping
    {
        public static int TextboxCount = 0;
        private int _textboxIndex;

        public TextboxMapping(ConversionContext ctx, int textboxIndex, ContentPart targetpart, XmlWriter writer)
            : base(ctx, targetpart, writer)
        {
            TextboxCount++;
            this._textboxIndex = textboxIndex;
        }

        public TextboxMapping(ConversionContext ctx, ContentPart targetpart, XmlWriter writer)
            : base(ctx, targetpart, writer)
        {
            TextboxCount++;
            this._textboxIndex = TextboxCount - 1;
        }

        public override void Apply(WordDocument doc)
        {
            this._doc = doc;

            this._writer.WriteStartElement("v", "textbox", OpenXmlNamespaces.VectorML);
            this._writer.WriteStartElement("w", "txbxContent", OpenXmlNamespaces.WordprocessingML);

            int cp = 0;
            int cpEnd = 0;
            BreakDescriptor bkd = null;
            int txtbxSubdocStart = doc.FIB.ccpText + doc.FIB.ccpFtn + doc.FIB.ccpHdr + doc.FIB.ccpAtn + doc.FIB.ccpEdn;

            if(this._targetPart.GetType() == typeof(MainDocumentPart))
            {
                cp = txtbxSubdocStart + doc.TextboxBreakPlex.CharacterPositions[this._textboxIndex];
                cpEnd = txtbxSubdocStart + doc.TextboxBreakPlex.CharacterPositions[this._textboxIndex + 1];
                bkd = (BreakDescriptor)doc.TextboxBreakPlex.Elements[this._textboxIndex];
            }
            if (this._targetPart.GetType() == typeof(HeaderPart) || this._targetPart.GetType() == typeof(FooterPart))
            {
                txtbxSubdocStart += doc.FIB.ccpTxbx;
                cp = txtbxSubdocStart + doc.TextboxBreakPlexHeader.CharacterPositions[this._textboxIndex];
                cpEnd = txtbxSubdocStart + doc.TextboxBreakPlexHeader.CharacterPositions[this._textboxIndex + 1];
                bkd = (BreakDescriptor)doc.TextboxBreakPlexHeader.Elements[this._textboxIndex];
            }

            //convert the textbox text
            this._lastValidPapx = this._doc.AllPapxFkps[0].grppapx[0];

            while (cp < cpEnd)
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
            this._writer.WriteEndElement();

            this._writer.Flush();
        }
    }
}
