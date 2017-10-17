using System;
using System.Collections.Generic;
using b2xtranslator.DocFileFormat;
using b2xtranslator.CommonTranslatorLib;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.Tools;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.WordprocessingMLMapping
{
    public abstract class DocumentMapping : 
        AbstractOpenXmlMapping,
        IMapping<WordDocument>
    {
        protected WordDocument _doc;
        protected ConversionContext _ctx;
        protected ParagraphPropertyExceptions _lastValidPapx;
        protected SectionPropertyExceptions _lastValidSepx;
        protected int _skipRuns = 0;
        protected int _sectionNr = 0;
        protected int _footnoteNr = 0;
        protected int _endnoteNr = 0;
        protected int _commentNr = 0;
        protected bool _writeInstrText = false;
        protected ContentPart _targetPart;

        private class Symbol
        {
            public string FontName;
            public string HexValue;
        }

        /// <summary>
        /// Creates a new DocumentMapping that writes to the given XmlWriter
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="targetPart"></param>
        public DocumentMapping(ConversionContext ctx, ContentPart targetPart, XmlWriter writer)
            : base(writer)
        {
            this._ctx = ctx;
            this._targetPart = targetPart;
            //Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
        }

        /// <summary>
        /// Creates a new DocumentMapping that creates a new XmLWriter on to the given ContentPart
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="targetPart"></param>
        public DocumentMapping(ConversionContext ctx, ContentPart targetPart)
            : base(XmlWriter.Create(targetPart.GetStream(), ctx.WriterSettings))
        {
            this._ctx = ctx;
            this._targetPart = targetPart;
            //Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
        }

        public abstract void Apply(WordDocument doc);

        #region TableConversion

        /// <summary>
        /// Writes the table starts at the given cp value
        /// </summary>
        /// <param name="cp">The cp at where the table begins</param>
        /// <returns>The character pointer to the first character after this table</returns>
        protected int writeTable(int initialCp, uint nestingLevel)
        {
            int cp = initialCp;
            int fc = this._doc.PieceTable.FileCharacterPositions[cp];
            var papx = findValidPapx(fc);
            var tai = new TableInfo(papx);

            //build the table grid
            var grid = buildTableGrid(cp, nestingLevel);

            //find first row end
            int fcRowEnd = findRowEndFc(cp, nestingLevel);
            var row1Tapx = new TablePropertyExceptions(findValidPapx(fcRowEnd), this._doc.DataStream);

            //start table
            this._writer.WriteStartElement("w", "tbl", OpenXmlNamespaces.WordprocessingML);

            //Convert it
            row1Tapx.Convert(new TablePropertiesMapping(this._writer, this._doc.Styles, grid));

            //convert all rows
            if (nestingLevel > 1)
            {
                //It's an inner table
                //only convert the cells with the given nesting level
                while (tai.iTap == nestingLevel)
                {
                    cp = writeTableRow(cp, grid, nestingLevel);
                    fc = this._doc.PieceTable.FileCharacterPositions[cp];
                    papx = findValidPapx(fc);
                    tai = new TableInfo(papx);
                }
            }
            else
            {
                //It's a outer table (nesting level 1)
                //convert until the end of table is reached
                while (tai.fInTable)
                {
                    cp = writeTableRow(cp, grid, nestingLevel);
                    fc = this._doc.PieceTable.FileCharacterPositions[cp];
                    papx = findValidPapx(fc);
                    tai = new TableInfo(papx);
                }
            }

            //close w:tbl
            this._writer.WriteEndElement();

            return cp;
        }

        /// <summary>
        /// Writes the table row that starts at the given cp value and ends at the next row end mark
        /// </summary>
        /// <param name="initialCp">The cp at where the row begins</param>
        /// <returns>The character pointer to the first character after this row</returns>
        protected int writeTableRow(int initialCp, List<short> grid, uint nestingLevel)
        {
            int cp = initialCp;
            int fc = this._doc.PieceTable.FileCharacterPositions[cp];
            var papx = findValidPapx(fc);
            var tai = new TableInfo(papx);

            //start w:tr
            this._writer.WriteStartElement("w", "tr", OpenXmlNamespaces.WordprocessingML);

            //convert the properties
            int fcRowEnd = findRowEndFc(cp, nestingLevel);
            var rowEndPapx = findValidPapx(fcRowEnd);
            var tapx = new TablePropertyExceptions(rowEndPapx, this._doc.DataStream);
            var chpxs = this._doc.GetCharacterPropertyExceptions(fcRowEnd, fcRowEnd + 1);

            if (tapx != null)
            {
                tapx.Convert(new TableRowPropertiesMapping(this._writer, chpxs[0]));
            }
            int gridIndex = 0;
            int cellIndex = 0;

            if (nestingLevel > 1)
            {
                //It's an inner table.
                //Write until the first "inner trailer paragraph" is reached
                while (!(this._doc.Text[cp] == TextMark.ParagraphEnd && tai.fInnerTtp) && tai.fInTable)
                {
                    cp = writeTableCell(cp, tapx, grid, ref gridIndex, cellIndex, nestingLevel);
                    cellIndex++;

                    //each cell has it's own PAPX
                    fc = this._doc.PieceTable.FileCharacterPositions[cp];
                    papx = findValidPapx(fc);
                    tai = new TableInfo(papx);
                }
            }
            else
            {
                //It's a outer table
                //Write until the first "row end trailer paragraph" is reached
                while (!(this._doc.Text[cp] == TextMark.CellOrRowMark && tai.fTtp) && tai.fInTable)
                {
                    cp = writeTableCell(cp, tapx, grid, ref gridIndex, cellIndex, nestingLevel);
                    cellIndex++;

                    //each cell has it's own PAPX
                    fc = this._doc.PieceTable.FileCharacterPositions[cp];
                    papx = findValidPapx(fc);
                    tai = new TableInfo(papx);
                }
            }


            //end w:tr
            this._writer.WriteEndElement();

            //skip the row end mark
            cp++;

            return cp;
        }


        /// <summary>
        /// Writes the table cell that starts at the given cp value and ends at the next cell end mark
        /// </summary>
        /// <param name="initialCp">The cp at where the cell begins</param>
        /// <param name="tapx">The TAPX that formats the row to which the cell belongs</param>
        /// <param name="gridIndex">The index of this cell in the grid</param>
        /// <param name="gridIndex">The grid</param>
        /// <returns>The character pointer to the first character after this cell</returns>
        protected int writeTableCell(int initialCp, TablePropertyExceptions tapx, List<short> grid, ref int gridIndex, int cellIndex, uint nestingLevel)
        {
            int cp = initialCp;

            //start w:tc
            this._writer.WriteStartElement("w", "tc", OpenXmlNamespaces.WordprocessingML);

            //find cell end
            int cpCellEnd = findCellEndCp(initialCp, nestingLevel);
            
            //convert the properties
            var mapping = new TableCellPropertiesMapping(this._writer, grid, gridIndex, cellIndex);
            if (tapx != null)
            {
                tapx.Convert(mapping);
            }
            gridIndex = gridIndex + mapping.GridSpan;
            

            //write the paragraphs of the cell
            while (cp < cpCellEnd)
            {
                //cp = writeParagraph(cp);
                int fc = this._doc.PieceTable.FileCharacterPositions[cp];
                var papx = findValidPapx(fc);
                var tai = new TableInfo(papx);

                //cp = writeParagraph(cp);

                if (tai.iTap > nestingLevel)
                {
                    //write the inner table if this is not a inner table (endless loop)
                    cp = writeTable(cp, tai.iTap);

                    //after a inner table must be at least one paragraph
                    //if (cp >= cpCellEnd)
                    //{
                    //    _writer.WriteStartElement("w", "p", OpenXmlNamespaces.WordprocessingML);
                    //    _writer.WriteEndElement();
                    //}
                }
                else
                {
                    //this PAPX is for a normal paragraph
                    cp = writeParagraph(cp);
                }
            }


            //end w:tc
            this._writer.WriteEndElement();

            return cp;
        }


        /// <summary>
        /// Builds a list that contains the width of the several columns of the table.
        /// </summary>
        /// <param name="initialCp"></param>
        /// <returns></returns>
        protected List<short> buildTableGrid(int initialCp, uint nestingLevel)
        {
            var backup = this._lastValidPapx;

            var boundaries = new List<short>();
            var grid = new List<short>();
            int cp = initialCp;
            int fc = this._doc.PieceTable.FileCharacterPositions[cp];
            var papx = findValidPapx(fc);
            var tai = new TableInfo(papx);

            int fcRowEnd = findRowEndFc(cp, out cp, nestingLevel);

            while (tai.fInTable)
            {
                //check all SPRMs of this TAPX
                foreach (var sprm in papx.grpprl)
                {
                    //find the tDef SPRM
                    if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmTDefTable)
                    {
                        byte itcMac = sprm.Arguments[0];
                        for (int i = 0; i < itcMac; i++)
                        {
                            short boundary1 = System.BitConverter.ToInt16(sprm.Arguments, 1 + (i * 2));
                            if (!boundaries.Contains(boundary1))
                                boundaries.Add(boundary1);

                            short boundary2 = System.BitConverter.ToInt16(sprm.Arguments, 1 + ((i + 1) * 2));
                            if (!boundaries.Contains(boundary2))
                                boundaries.Add(boundary2);
                        }
                    }
                }

                //get the next papx
                papx = findValidPapx(fcRowEnd);
                tai = new TableInfo(papx);
                fcRowEnd = findRowEndFc(cp, out cp, nestingLevel);
            }

            //build the grid based on the boundaries
            boundaries.Sort();
            for (int i = 0; i < boundaries.Count - 1; i++)
            {
                grid.Add((short)(boundaries[i + 1] - boundaries[i]));
            }

            this._lastValidPapx = backup;
            return grid;
        }

        /// <summary>
        /// Finds the FC of the next row end mark.
        /// </summary>
        /// <param name="initialCp">Some CP before the row end</param>
        /// <param name="rowEndCp">The CP of the next row end mark</param>
        /// <returns>The FC of the next row end mark</returns>
        protected int findRowEndFc(int initialCp, out int rowEndCp, uint nestingLevel)
        {
            int cp = initialCp;
            int fc = this._doc.PieceTable.FileCharacterPositions[cp];
            var papx = findValidPapx(fc);
            var tai = new TableInfo(papx);

            if (nestingLevel > 1)
            {
                //Its an inner table.
                //Search the "inner table trailer paragraph"
                while (tai.fInnerTtp == false && tai.fInTable == true)
                {
                    while (this._doc.Text[cp] != TextMark.ParagraphEnd)
                    {
                        cp++;
                    }
                    fc = this._doc.PieceTable.FileCharacterPositions[cp];
                    papx = findValidPapx(fc);
                    tai = new TableInfo(papx);
                    cp++;
                }
            }
            else 
            {
                //Its an outer table.
                //Search the "table trailer paragraph"
                while (tai.fTtp == false && tai.fInTable == true)
                {
                    while (this._doc.Text[cp] != TextMark.CellOrRowMark)
                    {
                        cp++;
                    }
                    fc = this._doc.PieceTable.FileCharacterPositions[cp];
                    papx = findValidPapx(fc);
                    tai = new TableInfo(papx);
                    cp++;
                }
            }

            rowEndCp = cp;
            return fc;
        }

        /// <summary>
        /// Finds the FC of the next row end mark.
        /// </summary>
        /// <param name="cp"></param>
        /// <returns></returns>
        protected int findRowEndFc(int initialCp, uint nestingLevel)
        {
            int cp = initialCp;
            int fc = this._doc.PieceTable.FileCharacterPositions[cp];
            var papx = findValidPapx(fc);
            var tai = new TableInfo(papx);

            if (nestingLevel > 1)
            {
                //Its an inner table.
                //Search the "inner table trailer paragraph"
                while (tai.fInnerTtp == false && tai.fInTable == true)
                {
                    while (this._doc.Text[cp] != TextMark.ParagraphEnd)
                    {
                        cp++;
                    }
                    fc = this._doc.PieceTable.FileCharacterPositions[cp];
                    papx = findValidPapx(fc);
                    tai = new TableInfo(papx);
                    cp++;
                }
            }
            else
            {
                //Its an outer table.
                //Search the "table trailer paragraph"
                while (tai.fTtp == false && tai.fInTable == true)
                {
                    while (this._doc.Text[cp] != TextMark.CellOrRowMark)
                    {
                        cp++;
                    }
                    fc = this._doc.PieceTable.FileCharacterPositions[cp];
                    papx = findValidPapx(fc);
                    tai = new TableInfo(papx);
                    cp++;
                }
            }

            return fc;
        }


        protected int findCellEndCp(int initialCp, uint nestingLevel)
        {
            int cpCellEnd = initialCp;

            if (nestingLevel > 1)
            {
                int fc = this._doc.PieceTable.FileCharacterPositions[initialCp];
                var papx = findValidPapx(fc);
                var tai = new TableInfo(papx);

                while (!tai.fInnerTableCell)
                {
                    cpCellEnd++;

                    fc = this._doc.PieceTable.FileCharacterPositions[cpCellEnd];
                    papx = findValidPapx(fc);
                    tai = new TableInfo(papx);
                }
                cpCellEnd++;
            }
            else 
            {
                while (this._doc.Text[cpCellEnd] != TextMark.CellOrRowMark)
                {
                    cpCellEnd++;
                }
                cpCellEnd++;
            }

            return cpCellEnd;
        }


        #endregion

        #region ParagraphRunConversion

        /// <summary>
        /// Writes a Paragraph that starts at the given cp and 
        /// ends at the next paragraph end mark or section end mark
        /// </summary>
        /// <param name="cp"></param>
        protected int writeParagraph(int cp) 
        {
            //search the paragraph end
            int cpParaEnd = cp;
            while (this._doc.Text[cpParaEnd] != TextMark.ParagraphEnd &&
                this._doc.Text[cpParaEnd] != TextMark.CellOrRowMark &&
                !(this._doc.Text[cpParaEnd] == TextMark.PageBreakOrSectionMark && isSectionEnd(cpParaEnd)))
            {
                cpParaEnd++;
            }

            if (this._doc.Text[cpParaEnd] == TextMark.PageBreakOrSectionMark)
            {
                //there is a page break OR section mark,
                //write the section only if it's a section mark
                bool sectionEnd = isSectionEnd(cpParaEnd);
                cpParaEnd++;
                return writeParagraph(cp, cpParaEnd, sectionEnd);
            }
            else
            {
                cpParaEnd++;
                return writeParagraph(cp, cpParaEnd, false);
            }
        }

        /// <summary>
        /// Writes a Paragraph that starts at the given cpStart and 
        /// ends at the given cpEnd
        /// </summary>
        /// <param name="cpStart"></param>
        /// <param name="cpEnd"></param>
        /// <param name="sectionEnd">Set if this paragraph is the last paragraph of a section</param>
        /// <returns></returns>
        protected int writeParagraph(int initialCp, int cpEnd, bool sectionEnd)
        {
            int cp = initialCp;
            int fc = this._doc.PieceTable.FileCharacterPositions[cp];
            int fcEnd = this._doc.PieceTable.FileCharacterPositions[cpEnd];
            var papx = findValidPapx(fc);

            //get all CHPX between these boundaries to determine the count of runs
            var chpxs = this._doc.GetCharacterPropertyExceptions(fc, fcEnd);
            var chpxFcs = this._doc.GetFileCharacterPositions(fc, fcEnd);
            chpxFcs.Add(fcEnd);

            //the last of these CHPX formats the paragraph end mark
            var paraEndChpx = chpxs[chpxs.Count-1];

            //start paragraph
            this._writer.WriteStartElement("w", "p", OpenXmlNamespaces.WordprocessingML);

            //check for section properties
            if (sectionEnd)
            {
                //this is the last paragraph of this section
                //write properties with section properties
                if (papx != null)
                {
                    papx.Convert(new ParagraphPropertiesMapping(this._writer, this._ctx, this._doc, paraEndChpx, findValidSepx(cpEnd), this._sectionNr));
                }
                this._sectionNr++;
            }
            else
            {
                //write properties
                if (papx != null)
                {
                    papx.Convert(new ParagraphPropertiesMapping(this._writer, this._ctx, this._doc, paraEndChpx));
                }
            }

            //write a runs for each CHPX
            for (int i = 0; i < chpxs.Count; i++)
            {
                //get the FC range for this run
                int fcChpxStart = chpxFcs[i];
                int fcChpxEnd = chpxFcs[i + 1];

                //it's the first chpx and it starts before the paragraph
                if (i == 0 && fcChpxStart < fc)
                {
                    //so use the FC of the paragraph
                    fcChpxStart = fc;
                }

                //it's the last chpx and it exceeds the paragraph
                if (i == (chpxs.Count - 1) && fcChpxEnd > fcEnd)
                {
                    //so use the FC of the paragraph
                    fcChpxEnd = fcEnd;
                }

                //read the chars that are formatted via this CHPX
                var chpxChars = this._doc.PieceTable.GetChars(fcChpxStart, fcChpxEnd, this._doc.WordDocumentStream);

                //search for bookmarks in the chars
                var bookmarks = searchBookmarks(chpxChars, cp);

                //if there are bookmarks in this run, split the run into several runs
                if (bookmarks.Count > 0)
                {
                    var runs = splitCharList(chpxChars, bookmarks);
                    for (int s = 0; s < runs.Count; s++)
                    {
                        if (this._doc.BookmarkStartPlex.CharacterPositions.Contains(cp) &&
                            this._doc.BookmarkEndPlex.CharacterPositions.Contains(cp))
                        {
                            //there start and end bookmarks here

                            //so get all bookmarks that end here
                            for (int b = 0; b < this._doc.BookmarkEndPlex.CharacterPositions.Count; b++)
                            {
                                if (this._doc.BookmarkEndPlex.CharacterPositions[b] == cp)
                                {
                                    //and check if the matching start bookmark also starts here
                                    if (this._doc.BookmarkStartPlex.CharacterPositions[b] == cp)
                                    {
                                        //then write a start and a end
                                        if (this._doc.BookmarkStartPlex.Elements.Count > b)
                                        {
                                            writeBookmarkStart((BookmarkFirst)this._doc.BookmarkStartPlex.Elements[b]);
                                            writeBookmarkEnd((BookmarkFirst)this._doc.BookmarkStartPlex.Elements[b]);
                                        }
                                    }
                                    else
                                    {
                                        //write a end
                                        writeBookmarkEnd((BookmarkFirst)this._doc.BookmarkStartPlex.Elements[b]);
                                    }
                                }
                            }

                            writeBookmarkStarts(cp);
                        }
                        else if (this._doc.BookmarkStartPlex.CharacterPositions.Contains(cp))
                        {
                            writeBookmarkStarts(cp);
                        }
                        else if (this._doc.BookmarkEndPlex.CharacterPositions.Contains(cp))
                        {
                            writeBookmarkEnds(cp);
                        }

                        cp = writeRun(runs[s], chpxs[i], cp);
                    }
                }
                else
                {
                    cp = writeRun(chpxChars, chpxs[i], cp);
                }
            }

            //end paragraph
            this._writer.WriteEndElement();

            return cpEnd++;
        }

        /// <summary>
        /// Writes a run with the given characters and CHPX
        /// </summary>
        protected int writeRun(List<char> chars, CharacterPropertyExceptions chpx, int initialCp)
        {
            int cp = initialCp;

            if (this._skipRuns <= 0 && chars.Count > 0)
            {
                var rev = new RevisionData(chpx);

                if (rev.Type == RevisionData.RevisionType.Deleted)
                {
                    //If it's a deleted run
                    this._writer.WriteStartElement("w", "del", OpenXmlNamespaces.WordprocessingML);
                    this._writer.WriteAttributeString("w", "author", OpenXmlNamespaces.WordprocessingML, "[b2x: could not retrieve author]");
                    this._writer.WriteAttributeString("w", "date", OpenXmlNamespaces.WordprocessingML, "[b2x: could not retrieve date]");
                }
                else if (rev.Type == RevisionData.RevisionType.Inserted)
                {
                    //if it's a inserted run
                    this._writer.WriteStartElement("w", "ins", OpenXmlNamespaces.WordprocessingML);
                    this._writer.WriteAttributeString("w", "author", OpenXmlNamespaces.WordprocessingML, this._doc.RevisionAuthorTable.Strings[rev.Isbt]);
                    rev.Dttm.Convert(new DateMapping(this._writer));
                }

                //start run
                this._writer.WriteStartElement("w", "r", OpenXmlNamespaces.WordprocessingML);

                //append rsids
                if (rev.Rsid != 0)
                {
                    string rsid = string.Format("{0:x8}", rev.Rsid);
                    this._writer.WriteAttributeString("w", "rsidR", OpenXmlNamespaces.WordprocessingML, rsid);
                    this._ctx.AddRsid(rsid);
                }
                if (rev.RsidDel != 0)
                {
                    string rsidDel = string.Format("{0:x8}", rev.RsidDel);
                    this._writer.WriteAttributeString("w", "rsidDel", OpenXmlNamespaces.WordprocessingML, rsidDel);
                    this._ctx.AddRsid(rsidDel);
                }
                if (rev.RsidProp != 0)
                {
                    string rsidProp = string.Format("{0:x8}", rev.RsidProp);
                    this._writer.WriteAttributeString("w", "rsidRPr", OpenXmlNamespaces.WordprocessingML, rsidProp);
                    this._ctx.AddRsid(rsidProp);
                }

                //convert properties
                chpx.Convert(new CharacterPropertiesMapping(this._writer, this._doc, rev, this._lastValidPapx, false));

                if (rev.Type == RevisionData.RevisionType.Deleted)
                    writeText(chars, cp, chpx, true);
                else
                    writeText(chars, cp, chpx, false);

                //end run
                this._writer.WriteEndElement();

                if (rev.Type == RevisionData.RevisionType.Deleted || rev.Type == RevisionData.RevisionType.Inserted)
                {
                    this._writer.WriteEndElement();
                }
            }
            else
            {
                this._skipRuns--;
            }

            return cp + chars.Count;
        }


        /// <summary>
        /// Writes the given text to the document
        /// </summary>
        /// <param name="chars"></param>
        protected void writeText(List<char> chars, int initialCp, CharacterPropertyExceptions chpx, bool writeDeletedText)
        {
            int cp = initialCp;
            bool fSpec = isSpecial(chpx);

            //detect text type
            string textType = "t";
            if(writeDeletedText)
                textType = "delText";
            else if(this._writeInstrText)
                textType = "instrText";
 
            //open a new w:t element
            writeTextStart(textType);

            //write text
            for (int i = 0; i < chars.Count; i++)
            {
                char c = chars[i];

                if (c == TextMark.Tab)
                {
                    this._writer.WriteEndElement();
                    this._writer.WriteElementString("w", "tab", OpenXmlNamespaces.WordprocessingML, "");
                    writeTextStart(textType);
                }
                else if (c == TextMark.HardLineBreak)
                {
                    //close previous w:t ...
                    this._writer.WriteEndElement();
                    this._writer.WriteElementString("w", "br", OpenXmlNamespaces.WordprocessingML, "");

                    writeTextStart(textType);
                }
                else if (c == TextMark.ParagraphEnd)
                {
                    //do nothing
                }
                else if (c == TextMark.PageBreakOrSectionMark)
                {
                    //write page break, section breaks are written by writeParagraph() method
                    if (!isSectionEnd(cp))
                    {
                        //close previous w:t ...
                        this._writer.WriteEndElement();

                        this._writer.WriteStartElement("w", "br", OpenXmlNamespaces.WordprocessingML);
                        this._writer.WriteAttributeString("w", "type", OpenXmlNamespaces.WordprocessingML, "page");
                        this._writer.WriteEndElement();

                        writeTextStart(textType);
                    }
                }
                else if (c == TextMark.ColumnBreak)
                {
                    //close previous w:t ...
                    this._writer.WriteEndElement();

                    this._writer.WriteStartElement("w", "br", OpenXmlNamespaces.WordprocessingML);
                    this._writer.WriteAttributeString("w", "type", OpenXmlNamespaces.WordprocessingML, "column");
                    this._writer.WriteEndElement();

                    writeTextStart(textType);
                }
                else if (c == TextMark.FieldBeginMark)
                {
                    //close previous w:t ...
                    this._writer.WriteEndElement();

                    int cpFieldStart = initialCp + i;
                    int cpFieldEnd = searchNextTextMark(this._doc.Text, cpFieldStart, TextMark.FieldEndMark);
                    var f = new Field(this._doc.Text.GetRange(cpFieldStart, cpFieldEnd - cpFieldStart + 1));

                    if(f.FieldCode.StartsWith(" FORM"))
                    {
                        this._writer.WriteStartElement("w", "fldChar", OpenXmlNamespaces.WordprocessingML);
                        this._writer.WriteAttributeString("w", "fldCharType", OpenXmlNamespaces.WordprocessingML, "begin");

                        int cpPic = searchNextTextMark(this._doc.Text, cpFieldStart, TextMark.Picture);
                        if (cpPic < cpFieldEnd)
                        {
                            int fcPic = this._doc.PieceTable.FileCharacterPositions[cpPic];
                            var chpxPic = this._doc.GetCharacterPropertyExceptions(fcPic, fcPic + 1)[0];
                            var npbd = new NilPicfAndBinData(chpxPic, this._doc.DataStream);
                            var ffdata = new FormFieldData(npbd.binData);
                            ffdata.Convert(new FormFieldDataMapping(this._writer));
                        }

                        this._writer.WriteEndElement();
                    }
                    else if (f.FieldCode.StartsWith(" EMBED") || f.FieldCode.StartsWith(" LINK"))
                    {
                        this._writer.WriteStartElement("w", "object", OpenXmlNamespaces.WordprocessingML);

                        int cpPic = searchNextTextMark(this._doc.Text, cpFieldStart, TextMark.Picture);
                        int cpFieldSep = searchNextTextMark(this._doc.Text, cpFieldStart, TextMark.FieldSeperator);

                        if (cpPic < cpFieldEnd)
                        {
                            int fcPic = this._doc.PieceTable.FileCharacterPositions[cpPic];
                            var chpxPic = this._doc.GetCharacterPropertyExceptions(fcPic, fcPic + 1)[0];
                            var pic = new PictureDescriptor(chpxPic, this._doc.DataStream);

                            //append the origin attributes
                            this._writer.WriteAttributeString("w", "dxaOrig", OpenXmlNamespaces.WordprocessingML, (pic.dxaGoal + pic.dxaOrigin).ToString());
                            this._writer.WriteAttributeString("w", "dyaOrig", OpenXmlNamespaces.WordprocessingML, (pic.dyaGoal + pic.dyaOrigin).ToString());

                            pic.Convert(new VMLPictureMapping(this._writer, this._targetPart, true));

                            if (cpFieldSep < cpFieldEnd)
                            {
                                int fcFieldSep = this._doc.PieceTable.FileCharacterPositions[cpFieldSep];
                                var chpxSep = this._doc.GetCharacterPropertyExceptions(fcFieldSep, fcFieldSep + 1)[0];
                                var ole = new OleObject(chpxSep, this._doc.Storage);
                                ole.Convert(new OleObjectMapping(this._writer, this._doc, this._targetPart, pic));
                            }
                        }

                        this._writer.WriteEndElement();

                        this._skipRuns = 4;
                    }
                    else
                    {
                        this._writer.WriteStartElement("w", "fldChar", OpenXmlNamespaces.WordprocessingML);
                        this._writer.WriteAttributeString("w", "fldCharType", OpenXmlNamespaces.WordprocessingML, "begin");
                        this._writer.WriteEndElement();
                    }

                    this._writeInstrText = true;

                    writeTextStart("instrText");
                }
                else if (c == TextMark.FieldSeperator)
                {
                    //close previous w:t ...
                    this._writer.WriteEndElement();

                    this._writer.WriteStartElement("w", "fldChar", OpenXmlNamespaces.WordprocessingML);
                    this._writer.WriteAttributeString("w", "fldCharType", OpenXmlNamespaces.WordprocessingML, "separate");
                    this._writer.WriteEndElement();

                    writeTextStart(textType);
                }
                else if (c == TextMark.FieldEndMark)
                {
                    //close previous w:t ...
                    this._writer.WriteEndElement();

                    this._writer.WriteStartElement("w", "fldChar", OpenXmlNamespaces.WordprocessingML);
                    this._writer.WriteAttributeString("w", "fldCharType", OpenXmlNamespaces.WordprocessingML, "end");
                    this._writer.WriteEndElement();

                    this._writeInstrText = false;

                    writeTextStart("t");
                }
                else if (c == TextMark.Symbol && fSpec)
                {
                    //close previous w:t ...
                    this._writer.WriteEndElement();

                    var s = getSymbol(chpx);
                    this._writer.WriteStartElement("w", "sym", OpenXmlNamespaces.WordprocessingML);
                    this._writer.WriteAttributeString("w", "font", OpenXmlNamespaces.WordprocessingML, s.FontName);
                    this._writer.WriteAttributeString("w", "char", OpenXmlNamespaces.WordprocessingML, s.HexValue);
                    this._writer.WriteEndElement();

                    writeTextStart(textType);
                }
                else if (c == TextMark.DrawnObject && fSpec)
                {
                    FileShapeAddress fspa = null;
                    if (GetType() == typeof(MainDocumentMapping))
                    {
                        fspa = (FileShapeAddress)this._doc.OfficeDrawingPlex.GetStruct(cp);
                    }
                    else if (GetType() == typeof(HeaderMapping) || GetType() == typeof(FooterMapping))
                    {
                        int headerCp = cp - this._doc.FIB.ccpText - this._doc.FIB.ccpFtn;
                        fspa = (FileShapeAddress)this._doc.OfficeDrawingPlexHeader.GetStruct(headerCp);
                    }
                    if (fspa != null)
                    {
                        var shape = this._doc.OfficeArtContent.GetShapeContainer(fspa.spid);
                        if (shape != null)
                        {
                            //close previous w:t ...
                            this._writer.WriteEndElement();
                            this._writer.WriteStartElement("w", "pict", OpenXmlNamespaces.WordprocessingML);

                            shape.Convert(new VMLShapeMapping(this._writer, this._targetPart, fspa, null, this._ctx));

                            this._writer.WriteEndElement();
                            writeTextStart(textType);
                        }
                    }
                }
                else if (c == TextMark.Picture && fSpec)
                {
                    var pict = new PictureDescriptor(chpx, this._doc.DataStream);
                    if (pict.mfp.mm > 98 && pict.ShapeContainer != null)
                    {
                        //close previous w:t ...
                        this._writer.WriteEndElement();
                        this._writer.WriteStartElement("w", "pict", OpenXmlNamespaces.WordprocessingML);

                        if (isWordArtShape(pict.ShapeContainer))
                        {
                            // a PICT without a BSE can stand for a WordArt Shape
                            pict.ShapeContainer.Convert(new VMLShapeMapping(this._writer, this._targetPart, null, pict, this._ctx));
                        }
                        else
                        {
                            // it's a normal picture
                            pict.Convert(new VMLPictureMapping(this._writer, this._targetPart, false));
                        }

                        this._writer.WriteEndElement();
                        writeTextStart(textType);
                    }                   
                }
                else if (c == TextMark.AutoNumberedFootnoteReference && fSpec)
                {
                    //close previous w:t ...
                    this._writer.WriteEndElement();

                    if (this.GetType() != typeof(FootnotesMapping) && this.GetType() != typeof(EndnotesMapping))
                    {
                        //it's in the document
                        if (this._doc.FootnoteReferencePlex.CharacterPositions.Contains(cp))
                        {
                            this._writer.WriteStartElement("w", "footnoteReference", OpenXmlNamespaces.WordprocessingML);
                            this._writer.WriteAttributeString("w", "id", OpenXmlNamespaces.WordprocessingML, this._footnoteNr.ToString());
                            this._writer.WriteEndElement();
                            this._footnoteNr++;
                        }
                        else if (this._doc.EndnoteReferencePlex.CharacterPositions.Contains(cp))
                        {
                            this._writer.WriteStartElement("w", "endnoteReference", OpenXmlNamespaces.WordprocessingML);
                            this._writer.WriteAttributeString("w", "id", OpenXmlNamespaces.WordprocessingML, this._endnoteNr.ToString());
                            this._writer.WriteEndElement();
                            this._endnoteNr++;
                        }
                    }
                    else
                    {
                        // it's not the document, write the short ref
                        if(this.GetType() != typeof(FootnotesMapping))
                        {
                            this._writer.WriteElementString("w", "footnoteRef", OpenXmlNamespaces.WordprocessingML, "");
                        }
                        if (this.GetType() != typeof(EndnotesMapping))
                        {
                            this._writer.WriteElementString("w", "endnoteRef", OpenXmlNamespaces.WordprocessingML, "");
                        }
                    }

                    writeTextStart(textType);
                }
                else if (c == TextMark.AnnotationReference)
                {
                    //close previous w:t ...
                    this._writer.WriteEndElement();

                    if (this.GetType() != typeof(CommentsMapping))
                    {
                        this._writer.WriteStartElement("w", "commentReference", OpenXmlNamespaces.WordprocessingML);
                        this._writer.WriteAttributeString("w", "id", OpenXmlNamespaces.WordprocessingML, this._commentNr.ToString());
                        this._writer.WriteEndElement();
                    }
                    else
                    {
                        this._writer.WriteElementString("w", "annotationRef", OpenXmlNamespaces.WordprocessingML, "");
                    }

                    this._commentNr++;

                    writeTextStart(textType);
                }
                else if ((int)c > 31 && (int)c != 0xFFFF)
                {
                    this._writer.WriteChars(new char[] { c }, 0, 1);
                }

                cp++;
            }

            //close w:t
            this._writer.WriteEndElement();
        }


        protected void writeTextStart(string textType)
        {
            this._writer.WriteStartElement("w", textType, OpenXmlNamespaces.WordprocessingML);
            this._writer.WriteAttributeString("xml", "space", "", "preserve");
        }

        /// <summary>
        /// Writes a bookmark start element at the given position
        /// </summary>
        /// <param name="cp"></param>
        protected void writeBookmarkStarts(int cp)
        {
            if (this._doc.BookmarkStartPlex.CharacterPositions.Count > 1)
            {
                for (int b = 0; b < this._doc.BookmarkStartPlex.CharacterPositions.Count; b++)
                {
                    if (this._doc.BookmarkStartPlex.CharacterPositions[b] == cp)
                    {
                        if (this._doc.BookmarkStartPlex.Elements.Count > b)
                        {
                            writeBookmarkStart((BookmarkFirst)this._doc.BookmarkStartPlex.Elements[b]);
                        }
                    }
                }
            }
        }

        protected void writeBookmarkStart(BookmarkFirst bookmark)
        {
            //write bookmark start
            this._writer.WriteStartElement("w", "bookmarkStart", OpenXmlNamespaces.WordprocessingML);
            this._writer.WriteAttributeString("w", "id", OpenXmlNamespaces.WordprocessingML, bookmark.ibkl.ToString());
            this._writer.WriteAttributeString("w", "name", OpenXmlNamespaces.WordprocessingML, this._doc.BookmarkNames.Strings[bookmark.ibkl]);
            this._writer.WriteEndElement();
        }

        /// <summary>
        /// Writes a bookmark end element at the given position
        /// </summary>
        /// <param name="cp"></param>
        protected void writeBookmarkEnds(int cp)
        {
            if (this._doc.BookmarkEndPlex.CharacterPositions.Count > 1)
            {
                //write all bookmark ends
                for (int b = 0; b < this._doc.BookmarkEndPlex.CharacterPositions.Count; b++)
                {
                    if (this._doc.BookmarkEndPlex.CharacterPositions[b] == cp)
                    {
                        writeBookmarkEnd((BookmarkFirst)this._doc.BookmarkStartPlex.Elements[b]);
                    }
                }     
            }
        }

        protected void writeBookmarkEnd(BookmarkFirst bookmark)
        {
            //write bookmark end
            this._writer.WriteStartElement("w", "bookmarkEnd", OpenXmlNamespaces.WordprocessingML);
            this._writer.WriteAttributeString("w", "id", OpenXmlNamespaces.WordprocessingML, bookmark.ibkl.ToString());
            this._writer.WriteEndElement();
        }

        #endregion

        #region HelpFunctions

        protected bool isWordArtShape(ShapeContainer shape)
        {
            bool result = false;
            var options = shape.ExtractOptions();
            foreach (var entry in options)
            {
                if (entry.pid == ShapeOptions.PropertyId.gtextUNICODE)
                {
                    result = true;
                    break;
                }
            }
            return result;
        }

        /// <summary>
        /// Splits a list of characters into several lists
        /// </summary>
        /// <returns></returns>
        protected List<List<char>> splitCharList(List<char> chars, List<int> splitIndices)
        {
            var ret = new List<List<char>>();

            int startIndex = 0;

            //add the parts
            for (int i = 0; i < splitIndices.Count; i++)
            {
                int cch = splitIndices[i] - startIndex;
                if (cch > 0)
                {
                    ret.Add(chars.GetRange(startIndex, cch));
                }
                startIndex += cch;
            }

            //add the last part
            ret.Add(chars.GetRange(startIndex, chars.Count-startIndex));

            return ret;
        }

        /// <summary>
        /// Searches for bookmarks in the list of characters.
        /// </summary>
        /// <param name="chars"></param>
        /// <returns>A List with all bookmarks indices in the given character list</returns>
        protected List<int> searchBookmarks(List<char> chars, int initialCp)
        {
            var ret = new List<int>();
            int cp = initialCp;
            for (int i = 0; i < chars.Count; i++)
            {
                if (this._doc.BookmarkStartPlex.CharacterPositions.Contains(cp) ||
                    this._doc.BookmarkEndPlex.CharacterPositions.Contains(cp))
                {
                    ret.Add(i);
                }
                cp++;
            }
            return ret;
        }

        /// <summary>
        /// Searches the given List for the next FieldEnd character.
        /// </summary>
        /// <param name="chars">The List of chars</param>
        /// <param name="initialCp">The position where the search should start</param>
        /// <param name="mark">The TextMark</param>
        /// <returns>The position of the next FieldEnd mark</returns>
        protected int searchNextTextMark(List<char> chars, int initialCp, char mark)
        {
            int ret = initialCp;
            for (int i = initialCp; i < chars.Count; i++)
            {
                if (chars[i] == mark)
                {
                    ret = i;
                    break;
                }
            }
            return ret;
        }

        /// <summary>
        /// Checks if the PAPX is old
        /// </summary>
        /// <param name="chpx">The PAPX</param>
        /// <returns></returns>
        protected bool isOld(ParagraphPropertyExceptions papx)
        {
            bool ret = false;
            foreach (var sprm in papx.grpprl)
            {
                if(sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPWall)
                {
                    //sHasOldProps
                    ret = true;
                    break;
                }
            }
            return ret;
        }

        /// <summary>
        /// Checks if the CHPX is special
        /// </summary>
        /// <param name="chpx">The CHPX</param>
        /// <returns></returns>
        protected bool isSpecial(CharacterPropertyExceptions chpx)
        {
            bool ret = false;
            foreach (var sprm in chpx.grpprl)
            {
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmCPicLocation ||
                    sprm.OpCode == SinglePropertyModifier.OperationCode.sprmCHsp)
                {
                    //special picture
                    ret = true;
                    break;
                }
                else if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmCSymbol)
                {
                    //special symbol
                    ret = true;
                    break;
                }
                else if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmCFSpec)
                {
                    //special value
                    ret = Utils.ByteToBool(sprm.Arguments[0]);
                    break;
                }
            }
            return ret;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="chpx"></param>
        /// <returns></returns>
        private Symbol getSymbol(CharacterPropertyExceptions chpx)
        {
            Symbol ret = null;
            foreach (var sprm in chpx.grpprl)
            {
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmCSymbol)
                {
                    //special symbol
                    ret = new Symbol();
                    short fontIndex = System.BitConverter.ToInt16(sprm.Arguments, 0);
                    short code = System.BitConverter.ToInt16(sprm.Arguments, 2);

                    var ffn = (FontFamilyName)this._doc.FontTable.Data[fontIndex];
                    ret.FontName = ffn.xszFtn;
                    ret.HexValue = string.Format("{0:x4}", code);
                    break;
                }
            }
            return ret;
        }

        /// <summary>
        /// Looks into the section table to find out if this CP is the end of a section
        /// </summary>
        /// <param name="cp"></param>
        /// <returns></returns>
        protected bool isSectionEnd(int cp)
        {
            bool result = false;

            //if cp is the last char of a section, the next section will start at cp +1
            int search = cp + 1;

            for (int i = 0; i < this._doc.SectionPlex.CharacterPositions.Count; i++)
            {
                if (this._doc.SectionPlex.CharacterPositions[i] == search)
                {
                    result = true;
                    break;
                }
            }

            return result;
        }

        /// <summary>
        /// Finds the PAPX that is valid for the given FC.
        /// </summary>
        /// <param name="fc"></param>
        /// <returns></returns>
        protected ParagraphPropertyExceptions findValidPapx(int fc)
        {
            ParagraphPropertyExceptions ret = null;

            if(this._doc.AllPapx.ContainsKey(fc))
            {
                ret = this._doc.AllPapx[fc];
                this._lastValidPapx = ret;
            }
            else
            {
                ret = this._lastValidPapx;
            }

            return ret;
        }

        /// <summary>
        /// Finds the SEPX that is valid for the given CP.
        /// </summary>
        /// <param name="cp"></param>
        /// <returns></returns>
        protected SectionPropertyExceptions findValidSepx(int cp)
        {
            SectionPropertyExceptions ret = null;

            try
            {
                ret = this._doc.AllSepx[cp];
                this._lastValidSepx = ret;
            }
            catch (KeyNotFoundException)
            {
                //there is no SEPX at this position, 
                //so the previous SEPX is valid for this cp

                int lastKey = this._doc.SectionPlex.CharacterPositions[1];
                foreach (int key in this._doc.AllSepx.Keys)
                {
                    if (cp > lastKey && cp < key)
                    {
                        ret = this._doc.AllSepx[lastKey];
                        break;
                    }
                    else
                    {
                        lastKey = key;
                    }
                }
            }

            return ret;
        }

        #endregion
    }
}
