using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.DocFileFormat;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.OpenXmlLib.WordprocessingML;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class SectionPropertiesMapping :
        PropertiesMapping,
        IMapping<SectionPropertyExceptions>
    {
        private XmlElement _sectPr;
        private int _sectNr;
        private ConversionContext _ctx;
       private SectionType _type = SectionType.nextPage;

        private enum SectionType
        {
            continuous = 0,
            nextColumn,
            nextPage,
            evenPage,
            oddPage
        }

        private enum PageOrientation
        {
            portrait = 1,
            landscape
        }

        private enum DocGridType
        {
            Default,
            lines,
            linesAndChars,
            snapToChars,
        }

        private enum FootnoteRestartCode
        {
            continuous,
            eachSect,
            eachPage
        }

        private enum PageNumberFormatCode
        {
            Decimal,
            upperRoman,
            lowerRoman,
            upperLetter,
            lowerLetter,
            ordinal,
            cardinalText,
            ordinalText,
            hex,
            chicago,
            ideographDigital,
            japaneseCounting,
            Aiueo,
            Iroha,
            decimalFullWidth,
            decimalHalfWidth,
            japaneseLegal,
            japaneseDigitalTenThousand,
            decimalEnclosedCircle,
            decimalFullWidth2,
            aiueoFullWidth,
            irohaFullWidth,
            decimalZero,
            bullet,
            ganada,
            chosung,
            decimalEnclosedFullstop,
            decimalEnclosedParen,
            decimalEnclosedCircleChinese,
            ideographEnclosedCircle,
            ideographTraditional,
            ideographZodiac,
            ideographZodiacTraditional,
            taiwaneseCounting,
            ideographLegalTraditional,
            taiwaneseCountingThousand,
            taiwaneseDigital,
            chineseCounting,
            chineseLegalSimplified,
            chineseCountingThousand,
            Decimal2,
            koreanDigital
        }

        private int _colNumber;
        private short[] _colSpace;
        private short[] _colWidth;
        private short _pgWidth, _marLeft, _marRight;

        /// <summary>
        /// Creates a new SectionPropertiesMapping which writes the 
        /// properties to the given writer
        /// </summary>
        /// <param name="writer">The XmlWriter</param>
        public SectionPropertiesMapping(XmlWriter writer, ConversionContext ctx, int sectionNr)
            : base(writer)
        {
            this._ctx = ctx;
            this._sectPr = this._nodeFactory.CreateElement("w", "sectPr", OpenXmlNamespaces.WordprocessingML);
            this._sectNr = sectionNr;
        }

        /// <summary>
        /// Creates a new SectionPropertiesMapping which appends 
        /// the properties to a given node.
        /// </summary>
        /// <param name="sectPr">The sectPr node</param>
        public SectionPropertiesMapping(XmlElement sectPr, ConversionContext ctx, int sectionNr) 
            : base(null)
        {
            this._ctx = ctx;
            this._nodeFactory = sectPr.OwnerDocument;
            this._sectPr = sectPr;
            this._sectNr = sectionNr;
        }

        /// <summary>
        /// Converts the given SectionPropertyExceptions
        /// </summary>
        /// <param name="sepx"></param>
        public void Apply(SectionPropertyExceptions sepx)
        {
            var pgMar = this._nodeFactory.CreateElement("w", "pgMar", OpenXmlNamespaces.WordprocessingML);
            var pgSz = this._nodeFactory.CreateElement("w", "pgSz", OpenXmlNamespaces.WordprocessingML);
            var docGrid = this._nodeFactory.CreateElement("w", "docGrid", OpenXmlNamespaces.WordprocessingML);
            var cols = this._nodeFactory.CreateElement("w", "cols", OpenXmlNamespaces.WordprocessingML);
            var pgBorders = this._nodeFactory.CreateElement("w", "pgBorders", OpenXmlNamespaces.WordprocessingML);
            var paperSrc = this._nodeFactory.CreateElement("w", "paperSrc", OpenXmlNamespaces.WordprocessingML);
            var footnotePr = this._nodeFactory.CreateElement("w", "footnotePr", OpenXmlNamespaces.WordprocessingML);
            var pgNumType = this._nodeFactory.CreateElement("w", "pgNumType", OpenXmlNamespaces.WordprocessingML);
            
            //convert headers of this section
            if (this._ctx.Doc.HeaderAndFooterTable.OddHeaders.Count > 0)
            {
                var evenHdr = this._ctx.Doc.HeaderAndFooterTable.EvenHeaders[this._sectNr];
                if (evenHdr != null)
                {
                    var evenPart = this._ctx.Docx.MainDocumentPart.AddHeaderPart();
                    this._ctx.Doc.Convert(new HeaderMapping(this._ctx, evenPart, evenHdr));
                    appendRef(this._sectPr, "headerReference", "even", evenPart.RelIdToString);
                }

                var oddHdr = this._ctx.Doc.HeaderAndFooterTable.OddHeaders[this._sectNr];
                if (oddHdr != null)
                {
                    var oddPart = this._ctx.Docx.MainDocumentPart.AddHeaderPart();
                    this._ctx.Doc.Convert(new HeaderMapping(this._ctx, oddPart, oddHdr));
                    appendRef(this._sectPr, "headerReference", "default", oddPart.RelIdToString);
                }

                var firstHdr = this._ctx.Doc.HeaderAndFooterTable.FirstHeaders[this._sectNr];
                if (firstHdr != null)
                {
                    var firstPart = this._ctx.Docx.MainDocumentPart.AddHeaderPart();
                    this._ctx.Doc.Convert(new HeaderMapping(this._ctx, firstPart, firstHdr));
                    appendRef(this._sectPr, "headerReference", "first", firstPart.RelIdToString);
                }
            }

            //convert footers of this section
            if (this._ctx.Doc.HeaderAndFooterTable.OddFooters.Count > 0)
            {
                var evenFtr = this._ctx.Doc.HeaderAndFooterTable.EvenFooters[this._sectNr];
                if (evenFtr != null)
                {
                    var evenPart = this._ctx.Docx.MainDocumentPart.AddFooterPart();
                    this._ctx.Doc.Convert(new FooterMapping(this._ctx, evenPart, evenFtr));
                    appendRef(this._sectPr, "footerReference", "even", evenPart.RelIdToString);
                }

                var oddFtr = this._ctx.Doc.HeaderAndFooterTable.OddFooters[this._sectNr];
                if (oddFtr != null)
                {
                    var oddPart = this._ctx.Docx.MainDocumentPart.AddFooterPart();
                    this._ctx.Doc.Convert(new FooterMapping(this._ctx, oddPart, oddFtr));
                    appendRef(this._sectPr, "footerReference", "default", oddPart.RelIdToString);
                }


                var firstFtr = this._ctx.Doc.HeaderAndFooterTable.FirstFooters[this._sectNr];
                if (firstFtr != null)
                {
                    var firstPart = this._ctx.Docx.MainDocumentPart.AddFooterPart();
                    this._ctx.Doc.Convert(new FooterMapping(this._ctx, firstPart, firstFtr));
                    appendRef(this._sectPr, "footerReference", "first", firstPart.RelIdToString);
                }
            }

            foreach (var sprm in sepx.grpprl)
            {
                switch (sprm.OpCode)
                {
                    //page margins
                    case SinglePropertyModifier.OperationCode.sprmSDxaLeft:
                        //left margin
                        this._marLeft = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        appendValueAttribute(pgMar, "left", this._marLeft.ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmSDxaRight:
                        //right margin
                        this._marRight = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        appendValueAttribute(pgMar, "right", this._marRight.ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmSDyaTop:
                        //top margin
                        appendValueAttribute(pgMar, "top", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmSDyaBottom:
                        //bottom margin
                        appendValueAttribute(pgMar, "bottom", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmSDzaGutter:
                        //gutter margin
                        appendValueAttribute(pgMar, "gutter", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmSDyaHdrTop:
                        //header margin
                        appendValueAttribute(pgMar, "header", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmSDyaHdrBottom:
                        //footer margin
                        appendValueAttribute(pgMar, "footer", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;

                    //page size and orientation
                    case SinglePropertyModifier.OperationCode.sprmSXaPage:
                        //width
                        this._pgWidth = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        appendValueAttribute(pgSz, "w", this._pgWidth.ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmSYaPage:
                        //height
                        appendValueAttribute(pgSz, "h", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmSBOrientation:
                        //orientation
                        appendValueAttribute(pgSz, "orient", ((PageOrientation)sprm.Arguments[0]).ToString());
                        break;

                    //paper source
                    case SinglePropertyModifier.OperationCode.sprmSDmBinFirst:
                        appendValueAttribute(paperSrc, "first", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmSDmBinOther:
                        appendValueAttribute(paperSrc, "other", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;

                    //page borders
                    case SinglePropertyModifier.OperationCode.sprmSBrcTop80:
                    case SinglePropertyModifier.OperationCode.sprmSBrcTop:
                        //top
                        var topBorder = this._nodeFactory.CreateElement("w", "top", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), topBorder);
                        addOrSetBorder(pgBorders, topBorder);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmSBrcLeft80:
                    case SinglePropertyModifier.OperationCode.sprmSBrcLeft:
                        //left
                        var leftBorder = this._nodeFactory.CreateElement("w", "left", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), leftBorder);
                        addOrSetBorder(pgBorders, leftBorder);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmSBrcBottom80:
                    case SinglePropertyModifier.OperationCode.sprmSBrcBottom:
                        //left
                        var bottomBorder = this._nodeFactory.CreateElement("w", "bottom", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), bottomBorder);
                        addOrSetBorder(pgBorders, bottomBorder);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmSBrcRight80:
                    case SinglePropertyModifier.OperationCode.sprmSBrcRight:
                        //left
                        var rightBorder = this._nodeFactory.CreateElement("w", "right", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), rightBorder);
                        addOrSetBorder(pgBorders, rightBorder);
                        break;

                    //footnote porperties
                    case SinglePropertyModifier.OperationCode.sprmSRncFtn:
                        //restart code
                        var fncFtn = FootnoteRestartCode.continuous;

                        //open office uses 1 byte values instead of 2 bytes values:
                        if (sprm.Arguments.Length == 2)
                        {
                            fncFtn = (FootnoteRestartCode)System.BitConverter.ToInt16(sprm.Arguments, 0);
                        }
                        if (sprm.Arguments.Length == 1)
                        {
                            fncFtn = (FootnoteRestartCode)sprm.Arguments[0];
                        }
                            
                        appendValueElement(footnotePr, "numRestart", fncFtn.ToString(), true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmSFpc:
                        //position code
                        short fpc = 0;
                        if(sprm.Arguments.Length == 2)
                            fpc = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        else
                            fpc = (short)sprm.Arguments[0];
                        if(fpc == 2)
                            appendValueElement(footnotePr, "pos", "beneathText", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmSNfcFtnRef:
                        //number format
                        short nfc = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        appendValueElement(footnotePr, "numFmt", NumberingMapping.GetNumberFormat(nfc), true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmSNFtn:
                        short nFtn = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        appendValueElement(footnotePr, "numStart", nFtn.ToString(), true);
                        break;

                    //doc grid
                    case SinglePropertyModifier.OperationCode.sprmSDyaLinePitch:
                        appendValueAttribute(docGrid, "linePitch", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmSDxtCharSpace:
                        appendValueAttribute(docGrid, "charSpace", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmSClm:
                        appendValueAttribute(docGrid, "type", ((DocGridType)System.BitConverter.ToInt16(sprm.Arguments, 0)).ToString());
                        break;

                    //columns
                    case SinglePropertyModifier.OperationCode.sprmSCcolumns:
                        this._colNumber = (int)(System.BitConverter.ToInt16(sprm.Arguments, 0) + 1);
                        this._colSpace = new short[this._colNumber];
                        appendValueAttribute(cols, "num", this._colNumber.ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmSDxaColumns:
                        //evenly spaced columns
                        appendValueAttribute(cols, "space", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmSDxaColWidth:
                        //there is at least one width set, so create the array
                        if(this._colWidth ==null)
                            this._colWidth = new short[this._colNumber];

                        byte index = sprm.Arguments[0];
                        short w = System.BitConverter.ToInt16(sprm.Arguments, 1);
                        this._colWidth[index] = w;
                        break;
                    case SinglePropertyModifier.OperationCode.sprmSDxaColSpacing:
                        //there is at least one space set, so create the array
                        if (this._colSpace == null)
                            this._colSpace = new short[this._colNumber];

                        this._colSpace[sprm.Arguments[0]] = System.BitConverter.ToInt16(sprm.Arguments, 1);
                        break;

                    //bidi
                    case SinglePropertyModifier.OperationCode.sprmSFBiDi:
                        appendFlagElement(this._sectPr, sprm, "bidi", true);
                        break;

                    //title page
                    case SinglePropertyModifier.OperationCode.sprmSFTitlePage:
                        appendFlagElement(this._sectPr, sprm, "titlePg", true);
                        break;

                    //RTL gutter
                    case SinglePropertyModifier.OperationCode.sprmSFRTLGutter:
                        appendFlagElement(this._sectPr, sprm, "rtlGutter", true);
                        break;

                    //type
                    case SinglePropertyModifier.OperationCode.sprmSBkc:
                        this._type = (SectionType)sprm.Arguments[0];
                        break;

                    //align
                    case SinglePropertyModifier.OperationCode.sprmSVjc:
                        appendValueElement(this._sectPr, "vAlign", sprm.Arguments[0].ToString(), true);
                        break;

                    //pgNumType
                    case SinglePropertyModifier.OperationCode.sprmSNfcPgn:
                        var pgnFc = (PageNumberFormatCode)sprm.Arguments[0];
                        appendValueAttribute(pgNumType, "fmt", pgnFc.ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmSPgnStart:
                        appendValueAttribute(pgNumType, "start", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                }
            }


            //build the columns
            if (this._colWidth != null)
            {
                //set to unequal width
                var equalWidth = this._nodeFactory.CreateAttribute("w", "equalWidth", OpenXmlNamespaces.WordprocessingML);
                equalWidth.Value = "0";
                cols.Attributes.Append(equalWidth);

                //calculate the width of the last column:
                //the last column width is not written to the document because it can be calculated.
                if (this._colWidth[this._colWidth.Length - 1] == 0)
                {
                    short lastColWidth = (short)(this._pgWidth - this._marLeft - this._marRight);
                    for (int i = 0; i < this._colWidth.Length - 1; i++)
                    {
                        lastColWidth -= this._colSpace[i];
                        lastColWidth -= this._colWidth[i];
                    }
                    this._colWidth[this._colWidth.Length - 1] = lastColWidth;
                }

                //append the xml elements
                for (int i = 0; i < this._colWidth.Length; i++)
                {
                    var col = this._nodeFactory.CreateElement("w", "col", OpenXmlNamespaces.WordprocessingML);
                    var w = this._nodeFactory.CreateAttribute("w", "w", OpenXmlNamespaces.WordprocessingML);
                    var space = this._nodeFactory.CreateAttribute("w", "space", OpenXmlNamespaces.WordprocessingML);
                    w.Value = this._colWidth[i].ToString();
                    space.Value = this._colSpace[i].ToString();
                    col.Attributes.Append(w);
                    col.Attributes.Append(space);
                    cols.AppendChild(col);
                }
            }

            //append the section type
            appendValueElement(this._sectPr, "type", this._type.ToString(), true);

            //append footnote properties
            if (footnotePr.ChildNodes.Count > 0)
            {
                this._sectPr.AppendChild(footnotePr);
            }

            //append page size
            if (pgSz.Attributes.Count > 0)
            {
                this._sectPr.AppendChild(pgSz);
            }

            //append borders
            if (pgBorders.ChildNodes.Count > 0)
            {
                this._sectPr.AppendChild(pgBorders);
            }

            //append margin
            if (pgMar.Attributes.Count > 0)
            {
                this._sectPr.AppendChild(pgMar);
            }

            //append paper info
            if (paperSrc.Attributes.Count > 0)
            {
                this._sectPr.AppendChild(paperSrc);
            }

            //append columns

            if (cols.Attributes.Count > 0 || cols.ChildNodes.Count > 0)
            {
                this._sectPr.AppendChild(cols);
            }

            //append doc grid
            if (docGrid.Attributes.Count > 0)
            {
                this._sectPr.AppendChild(docGrid);
            }

            //numType
            if (pgNumType.Attributes.Count > 0)
            {
                this._sectPr.AppendChild(pgNumType);
            }

            if (this._writer != null)
            {
                //write the properties
                this._sectPr.WriteTo(this._writer);
            }
        }

        private void appendRef(XmlElement parent, string element, string refType, string refId)
        {
            var headerRef = this._nodeFactory.CreateElement("w", element, OpenXmlNamespaces.WordprocessingML);
            
            var headerRefType = this._nodeFactory.CreateAttribute("w", "type", OpenXmlNamespaces.WordprocessingML);
            headerRefType.Value = refType;
            headerRef.Attributes.Append(headerRefType);

            var headerRefId = this._nodeFactory.CreateAttribute("r", "id", OpenXmlNamespaces.Relationships);
            headerRefId.Value = refId;
            headerRef.Attributes.Append(headerRefId);

            parent.AppendChild(headerRef);
        }
    }
}
