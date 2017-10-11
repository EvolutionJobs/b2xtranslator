

using System;
using System.Collections.Generic;
using b2xtranslator.PptFileFormat;
using b2xtranslator.CommonTranslatorLib;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PresentationMLMapping
{
    class TextMapping :
        AbstractOpenXmlMapping,
        IMapping<ClientTextbox>
    {
        protected ConversionContext _ctx;
        private string lang = "";
        private string altLang = "";
        private ShapeTreeMapping parentShapeTreeMapping = null;

        public TextMapping(ConversionContext ctx, XmlWriter writer)
            : base(writer)
        {
            this._ctx = ctx;
        }

        /// <summary>
        /// Returns the ParagraphRun of the given style that is active at the given index.
        /// </summary>
        /// <param name="style">style to use</param>
        /// <param name="forIdx">index to use</param>
        /// <returns>ParagraphRun or null in case no run was found</returns>
        protected static ParagraphRun GetParagraphRun(TextStyleAtom style, uint forIdx, ref int runCount)
        {
            if (style == null)
                return null;

            uint idx = 0;
            runCount = 0;

            foreach (var p in style.PRuns)
            {
                if (forIdx < idx + p.Length)
                    return p;

                idx += p.Length;
                runCount++;
            }

            return null;
        }

        /// <summary>
        /// Returns the MasterTextParagraphRun of the given style that is active at the given index.
        /// </summary>
        /// <param name="style">style to use</param>
        /// <param name="forIdx">index to use</param>
        /// <returns>ParagraphRun or null in case no run was found</returns>
        protected static MasterTextPropRun GetMasterTextPropRun(MasterTextPropAtom style, uint forIdx)
        {
            if (style == null)
                return new MasterTextPropRun();

            uint idx = 0;

            foreach (var p in style.MasterTextPropRuns)
            {
                if (forIdx < idx + p.count)
                    return p;

                idx += p.count;
            }

            return new MasterTextPropRun();
        }

        /// <summary>
        /// Returns the CharacterRun of the given style that is active at the given index.
        /// </summary>
        /// <param name="style">style to use</param>
        /// <param name="forIdx">index to use</param>
        /// <returns>CharacterRun or null in case no run was found</returns>
        protected static CharacterRun GetCharacterRun(TextStyleAtom style, uint forIdx)
        {
            if (style == null)
                return null;

            uint idx = 0;

            foreach (var c in style.CRuns)
            {
                if (forIdx < idx + c.Length)
                    return c;

                idx += c.Length;
            }

            return null;
        }

        protected static uint GetCharacterRunStart(TextStyleAtom style, uint forIdx)
        {
            if (style == null)
                return 0;

            uint idx = 0;

            foreach (var c in style.CRuns)
            {
                if (forIdx < idx + c.Length)
                    return idx;

                idx += c.Length;
            }

            return 0;
        }

        public void Apply(ClientTextbox textbox)
        {
            Apply(null, textbox, "", "", "", false);
        }

        public void Apply(ShapeTreeMapping pparentShapeTreeMapping, ClientTextbox textbox, string footertext, string headertext, string datetext, bool insideTable)
        {
            this.parentShapeTreeMapping = pparentShapeTreeMapping;
            var ms = new System.IO.MemoryStream(textbox.Bytes);
            var rec = Record.ReadRecord(ms);
            TextHeaderAtom thAtom = null;
            TextStyleAtom style = null;
            FooterMCAtom mca = null;
            TextSpecialInfoAtom sia = null;
            TextSpecialInfoAtom siaDefaults = null;
            TextRulerAtom ruler = null;
            var mciics = new List<MouseClickInteractiveInfoContainer>();
            MasterTextPropAtom masterTextProp = null;
            string text = "";
            string origText = "";
            var so = textbox.FirstAncestorWithType<ShapeContainer>().FirstChildWithType<ShapeOptions>();
            TextMasterStyleAtom defaultStyle = null;
            int lvl = 0;

            var parentSlide = textbox.FirstAncestorWithType<Slide>();
            if (parentSlide != null)
            foreach (var container in this._ctx.Ppt.DocumentRecord.AllChildrenWithType<SlideListWithText>())
            {
                if (container.Instance == 0)
                {
                    if (container.SlideToPlaceholderSpecialInfo.ContainsKey(parentSlide.PersistAtom))
                    {
                        siaDefaults = container.SlideToPlaceholderSpecialInfo[parentSlide.PersistAtom][0];
                    }
                }
            }

            switch (rec.TypeCode)
            {
                case 3999:
                    thAtom = (TextHeaderAtom)rec;
                    while (ms.Position < ms.Length)
                    {
                        rec = Record.ReadRecord(ms);

                        switch (rec.TypeCode)
                        {
                            case 0xfa0: //TextCharsAtom
                                text = ((TextCharsAtom)rec).Text;
                                origText = text;
                                thAtom.TextAtom = (TextAtom)rec;
                                break;
                            case 0xfa1: //TextRunStyleAtom
                                style = (TextStyleAtom)rec;
                                style.TextHeaderAtom = thAtom;
                                break;
                            case 0xfa6: //TextRulerAtom
                                ruler = (TextRulerAtom)rec;
                                break;
                            case 0xfa8: //TextBytesAtom
                                text = ((TextBytesAtom)rec).Text;
                                origText = text;
                                thAtom.TextAtom = (TextAtom)rec;
                                break;
                            case 0xfaa: //TextSpecialInfoAtom
                                sia = (TextSpecialInfoAtom)rec;
                                break;
                            case 0xfa2: //MasterTextPropAtom
                                masterTextProp = (MasterTextPropAtom)rec;
                                break;
                            case 0xfd8: //SlideNumberMCAtom
                                var snmca = (SlideNumberMCAtom)rec;

                                this._writer.WriteStartElement("a", "p", OpenXmlNamespaces.DrawingML);

                                this._writer.WriteStartElement("a", "fld", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteAttributeString("id", "{18109A10-03E4-4BE3-B6BB-0FCEF851AF87}");
                                this._writer.WriteAttributeString("type", "slidenum");
                                this._writer.WriteElementString("a", "t", OpenXmlNamespaces.DrawingML, "<#>");
                                this._writer.WriteEndElement(); //fld                                
                                this._writer.WriteStartElement("a", "endParaRPr", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteEndElement(); //endParaRPr
                                this._writer.WriteEndElement(); //p

                                text = text.Replace(origText.Substring(snmca.Position, 1), "");

                                break;
                            case 0xff2: //MouseClickInteractiveInfoContainer
                                mciics.Add((MouseClickInteractiveInfoContainer)rec);
                                break;
                            case 0xfdf: //MouseClickTextInteractiveInfoAtom
                                mciics[mciics.Count-1].Range = (MouseClickTextInteractiveInfoAtom)rec;
                                break;
                            case 0xff7: //DateTimeMCAtom
                                var d = (DateTimeMCAtom)rec;
                                string date = System.DateTime.Now.ToString();

                                //_writer.WriteStartElement("a", "p", OpenXmlNamespaces.DrawingML);

                                int runCount = 0;
                                var p = GetParagraphRun(style, 0, ref runCount);
                                var tp = GetMasterTextPropRun(masterTextProp, 0);
                                writeP(p, tp, so, ruler, defaultStyle,0);

                                this._writer.WriteStartElement("a", "fld", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteAttributeString("id", "{1023E2E8-AA53-4FEA-8F5C-1FABD68F61AB}");
                                this._writer.WriteAttributeString("type", "datetime1");

                                var r = GetCharacterRun(style, 0);
                                if (r != null)
                                {
                                    string dummy = "";
                                    string dummy2 = "";
                                    string dummy3 = "";
                                    RegularContainer slide = textbox.FirstAncestorWithType<Slide>();
                                    if (slide == null) slide = textbox.FirstAncestorWithType<Note>();
                                    if (slide == null) slide = textbox.FirstAncestorWithType<Handout>();
                                    new CharacterRunPropsMapping(this._ctx, this._writer).Apply(r, "rPr", slide, ref dummy, ref dummy2, ref dummy3, this.lang, this.altLang, defaultStyle,lvl,mciics,pparentShapeTreeMapping,0, insideTable);
                                }

                                this._writer.WriteElementString("a", "t", OpenXmlNamespaces.DrawingML, date);
                                this._writer.WriteEndElement(); //fld                                
                                this._writer.WriteStartElement("a", "endParaRPr", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteEndElement(); //endParaRPr
                                this._writer.WriteEndElement(); //p

                                text = text.Replace(origText.Substring(d.Position, 1), "");

                                foreach (var run in style.CRuns)
                                {
                                    run.Length += (uint)text.Length;
                                }
                                break;
                            case 0xff9: //HeaderMCAtom
                                var hmca = (HeaderMCAtom)rec;
                                text = text.Replace(origText.Substring(hmca.Position, 1), headertext);

                                foreach (var run in style.CRuns)
                                {
                                    run.Length += (uint)text.Length;
                                }
                                break;
                            case 0xffa: //FooterMCAtom
                                mca = (FooterMCAtom)rec;
                                text = text.Replace(origText.Substring(mca.Position, 1), footertext);

                                foreach (var run in style.CRuns)
                                {
                                    run.Length += (uint)text.Length;
                                }
                                break;
                            case 0xff8: //GenericDateMCAtom
                                var gdmca = (GenericDateMCAtom)rec;
                                text = text.Replace(origText.Substring(gdmca.Position, 1), datetext);

                                foreach (var run in style.CRuns)
                                {
                                    run.Length += (uint)text.Length;
                                }
                                break;
                            default:
                                //TextAtom textAtom = thAtom.TextAtom;
                                //text = (textAtom == null) ? "" : textAtom.Text;
                                //style = thAtom.TextStyleAtom;
                                break;
                        }
                    }
                    break;
                case 3998:
                    var otrAtom = (OutlineTextRefAtom)rec;
                    var slideListWithText = this._ctx.Ppt.DocumentRecord.RegularSlideListWithText;
                                  
                    var thAtoms = slideListWithText.SlideToPlaceholderTextHeaders[textbox.FirstAncestorWithType<Slide>().PersistAtom];
                    thAtom = thAtoms[otrAtom.Index];

                    if (thAtom.TextAtom != null) text = thAtom.TextAtom.Text;
                    if (thAtom.TextStyleAtom != null) style = thAtom.TextStyleAtom;

                    while (ms.Position < ms.Length)
                    {
                        rec = Record.ReadRecord(ms);
                        switch (rec.TypeCode)
                        {
                            case 0xfa6: //TextRulerAtom
                                ruler = (TextRulerAtom)rec;
                                break;
                            default:
                                break;
                        }
                    }

                    break;
                default:
                    throw new NotSupportedException("Can't find text for ClientTextbox without TextHeaderAtom and OutlineTextRefAtom");
            }

            uint idx = 0;                      

            var s = textbox.FirstAncestorWithType<Slide>();
           
            if (s != null)
            {
                try
                {
                    var a = s.FirstChildWithType<SlideAtom>();
                    if (Tools.Utils.BitmaskToBool(a.Flags, 0x01 << 1) && a.MasterId > 0)
                    {
                        var m = this._ctx.Ppt.FindMasterRecordById(a.MasterId);
                        foreach (var at in m.AllChildrenWithType<TextMasterStyleAtom>())
                        {
                            if (at.Instance == (int)thAtom.TextType)
                            {
                                defaultStyle = at;
                            }
                        }
                    }
                }
                catch (Exception)
                {                    
                    throw;
                }
                
            }

            //combine sia and siaDefaults
            var lstSIRuns = new Dictionary<TextSIRun, uint>();
            uint pos = 0;
            if (siaDefaults != null)
            foreach (var sirun in siaDefaults.Runs)
            {
                lstSIRuns.Add(sirun,pos);
                pos += sirun.count;
            }
            pos = 0;
            if (sia != null)
            foreach (var sirun in sia.Runs)
            {
                lstSIRuns.Add(sirun, pos);
                pos += sirun.count;
            }
            
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.hspMaster))
            {
                uint MasterID = so.OptionsByID[ShapeOptions.PropertyId.hspMaster].op;
            }

            if (text.Length == 0)
            {
                this._writer.WriteStartElement("a", "p", OpenXmlNamespaces.DrawingML);
                this._writer.WriteStartElement("a", "endParaRPr", OpenXmlNamespaces.DrawingML);
                // TODO...
                this._writer.WriteEndElement();
                this._writer.WriteEndElement();
            }
            else
            {
                var parlines = text.Split(new char[] { '\r' }); //text.Split(new char[] { '\v', '\r' });
                int internalOffset = 0;
                foreach (string parline in parlines)
                {
                    var runlines = parline.Split(new char[] { '\v' });

                    //each parline forms a paragraph
                    //each runline forms a run
                    
                    int runCount = 0;
                    var p = GetParagraphRun(style, idx, ref runCount);
                    var tp = GetMasterTextPropRun(masterTextProp, idx);

                    if (p != null) lvl = p.IndentLevel;

                    string runText;

                    writeP(p, tp, so, ruler, defaultStyle, runCount);

                    uint CharacterRunStart;
                    int len;

                    bool first = true;
                    bool textwritten = false;
                    foreach (string line in runlines)
                    {
                        uint offset = idx;

                        if (!first)
                        {
                            this._writer.WriteStartElement("a", "br", OpenXmlNamespaces.DrawingML);
                            var r = GetCharacterRun(style, idx + (uint)internalOffset + 1);
                            if (r != null)
                            {
                                string dummy = "";
                                string dummy2 = "";
                                string dummy3 = "";
                                RegularContainer slide = textbox.FirstAncestorWithType<Slide>();
                                if (slide == null) slide = textbox.FirstAncestorWithType<Note>();
                                if (slide == null) slide = textbox.FirstAncestorWithType<Handout>();
                                new CharacterRunPropsMapping(this._ctx, this._writer).Apply(r, "rPr", slide, ref dummy, ref dummy2, ref dummy3, this.lang, this.altLang, defaultStyle,lvl,mciics,pparentShapeTreeMapping,idx, insideTable);
                            }

                            this._writer.WriteEndElement();
                            if (line.Length == 0)
                            {
                               idx++;
                               internalOffset -= 1;
                            }
                            
                            internalOffset += 1;
                            
                        }

                        

                        while (idx < offset + line.Length)
                        {
                            textwritten = true;
                            len = line.Length;
                            CharacterRun r = null;
                            if (idx + (uint)internalOffset == 0)
                            {
                                r = GetCharacterRun(style, 0);
                                CharacterRunStart = GetCharacterRunStart(style, 0);
                            }
                            else
                            {
                                r = GetCharacterRun(style, idx + (uint)internalOffset);
                                CharacterRunStart = GetCharacterRunStart(style, idx + (uint)internalOffset);
                            }

                            if (r != null)
                            {                                
                                len = (int)(CharacterRunStart + r.Length - idx - internalOffset);
                                if (len > line.Length - idx + offset) len = (int)(line.Length - idx + offset);
                                if (len < 0) len = (int)(line.Length - idx + offset);
                                runText = line.Substring((int)(idx - offset), len);
                            }
                            else
                            {
                                runText = line.Substring((int)(idx - offset));
                                len = runText.Length;
                            }           
                            
                             //split runlines that partly contain a link
                            foreach (var mccic in mciics)
                            {
                                if (mccic.Range.begin <= idx + internalOffset && mccic.Range.end > idx + internalOffset)
                                {
                                    //link end before text ends
                                    if (mccic.Range.end < idx + internalOffset + len)
                                    {
                                        runText = line.Substring((int)(idx - offset),(int)(mccic.Range.end - mccic.Range.begin));
                                    }
                                }
                                else if (mccic.Range.begin >= idx + internalOffset && mccic.Range.end <= idx + internalOffset + len)
                                {
                                    //link starts inside the text
                                    runText = line.Substring((int)(idx - offset), (int)(mccic.Range.begin - (idx)));
                                }
                            }

                            //split runlines that contain a change of language
                            this.lang = "";
                            this.altLang = "";
                            foreach (var sirun in lstSIRuns.Keys)
                            {
                                uint start = lstSIRuns[sirun];
                                uint end = start + sirun.count; //languageRuns[start];

                                if (idx >= start && idx < end)
                                {
                                    if (end < idx + runText.Length)
                                    {
                                        runText = line.Substring((int)(idx - offset), (int)(end - idx));
                                    }

                                    if (sirun.si.lang)
                                    {
                                        switch (sirun.si.lid)
                                        {
                                            case 0x0: // no language
                                                break;
                                            case 0x13: //Any Dutch language is preferred over non-Dutch languages when proofing the text
                                                break;
                                            case 0x400: //no proofing
                                                break;
                                            default:
                                                try
                                                {
                                                    this.lang = System.Globalization.CultureInfo.GetCultureInfo(sirun.si.lid).IetfLanguageTag;
                                                }
                                                catch (Exception)
                                                {
                                                    //ignore
                                                }
                                                break;
                                        }
                                    }
                                    if (sirun.si.altLang)
                                    {
                                        switch (sirun.si.altLid)
                                        {
                                            case 0x0: // no language
                                                break;
                                            case 0x13: //Any Dutch language is preferred over non-Dutch languages when proofing the text
                                                break;
                                            case 0x400: //no proofing
                                                break;
                                            default:
                                                try
                                                {
                                                    this.altLang = System.Globalization.CultureInfo.GetCultureInfo(sirun.si.altLid).IetfLanguageTag;
                                                }
                                                catch (Exception)
                                                {
                                                    //ignore
                                                }
                                                break;
                                        }
                                    }
                                                         
                                }
                            }


                            this._writer.WriteStartElement("a", "r", OpenXmlNamespaces.DrawingML);

                            string dummy = "";
                            string dummy2 = "";
                            string dummy3 = "";
                            RegularContainer slide = textbox.FirstAncestorWithType<Slide>();
                            if (slide == null) slide = textbox.FirstAncestorWithType<Note>();
                            if (slide == null) slide = textbox.FirstAncestorWithType<Handout>();

                            if (r != null || defaultStyle != null)
                            {
                                new CharacterRunPropsMapping(this._ctx, this._writer).Apply(r, "rPr", slide,ref dummy, ref dummy2, ref dummy3, this.lang, this.altLang, defaultStyle,lvl, mciics, pparentShapeTreeMapping, idx, insideTable);
                            }
                            else
                            {
                                this._writer.WriteStartElement("a", "rPr", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteAttributeString("lang", this.lang);
                                this._writer.WriteEndElement(); 
                            }

                            this._writer.WriteStartElement("a", "t", OpenXmlNamespaces.DrawingML);

                            this._writer.WriteValue(runText.Replace(Char.ToString((char)0x05), ""));

                            this._writer.WriteEndElement();

                            this._writer.WriteEndElement();

                            idx += (uint)runText.Length; // +1;
                           
                        }

                        first = false;

                       
                    }

                    if (!textwritten)
                    {
                        var r = GetCharacterRun(style, idx + (uint)internalOffset);
                        if (r != null)
                        {
                            string dummy = "";
                            string dummy2 = "";
                            string dummy3 = "";
                            RegularContainer slide = textbox.FirstAncestorWithType<Slide>();
                            if (slide == null) slide = textbox.FirstAncestorWithType<Note>();
                            if (slide == null) slide = textbox.FirstAncestorWithType<Handout>();
                            new CharacterRunPropsMapping(this._ctx, this._writer).Apply(r, "endParaRPr", slide, ref dummy, ref dummy2, ref dummy3, this.lang, this.altLang, defaultStyle,lvl,mciics,pparentShapeTreeMapping,idx, insideTable);
                        }

                    }
                    else
                    {
                        this._writer.WriteStartElement("a", "endParaRPr", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteEndElement();
                    }

                    this._writer.WriteEndElement();

                    idx += 1;
                }

            }

        }

        private void writeP(ParagraphRun p, MasterTextPropRun tp, ShapeOptions so, TextRulerAtom ruler, TextMasterStyleAtom defaultStyle, int runCount)
        {
            int writtenLeftMargin = -1;
            this._writer.WriteStartElement("a", "p", OpenXmlNamespaces.DrawingML);

            this._writer.WriteStartElement("a", "pPr", OpenXmlNamespaces.DrawingML);
            if (p == null)
            {
                this._writer.WriteAttributeString("lvl", tp.indentLevel.ToString());
                               

                if (defaultStyle != null && defaultStyle.PRuns.Count > tp.indentLevel)
                {
                    if (defaultStyle.PRuns[tp.indentLevel].LeftMarginPresent)
                    {
                        this._writer.WriteAttributeString("marL", Utils.MasterCoordToEMU((int)defaultStyle.PRuns[tp.indentLevel].LeftMargin).ToString());
                        writtenLeftMargin = (int)defaultStyle.PRuns[tp.indentLevel].LeftMargin;
                    }
                    if (defaultStyle.PRuns[tp.indentLevel].IndentPresent)
                    {
                        this._writer.WriteAttributeString("indent", (-1 * (Utils.MasterCoordToEMU((int)(defaultStyle.PRuns[tp.indentLevel].LeftMargin - defaultStyle.PRuns[tp.indentLevel].Indent)))).ToString());
                    }


                    if (defaultStyle.PRuns[tp.indentLevel].AlignmentPresent)
                    {
                        switch (defaultStyle.PRuns[tp.indentLevel].Alignment)
                        {
                            case 0x0000: //Left
                                this._writer.WriteAttributeString("algn", "l");
                                break;
                            case 0x0001: //Center
                                this._writer.WriteAttributeString("algn", "ctr");
                                break;
                            case 0x0002: //Right
                                this._writer.WriteAttributeString("algn", "r");
                                break;
                            case 0x0003: //Justify
                                this._writer.WriteAttributeString("algn", "just");
                                break;
                            case 0x0004: //Distributed
                                this._writer.WriteAttributeString("algn", "dist");
                                break;
                            case 0x0005: //ThaiDistributed
                                this._writer.WriteAttributeString("algn", "thaiDist");
                                break;
                            case 0x0006: //JustifyLow
                                this._writer.WriteAttributeString("algn", "justLow");
                                break;
                        }
                    }

                    if (defaultStyle.PRuns[tp.indentLevel].BulletFlagsFieldPresent)
                    {
                        if ((defaultStyle.PRuns[tp.indentLevel].BulletFlags & (ushort)ParagraphMask.HasBullet) == 0)
                        {
                            this._writer.WriteElementString("a", "buNone", OpenXmlNamespaces.DrawingML, "");
                        }
                        else
                        {
                            //if (defaultStyle.PRuns[tp.indentLevel].BulletColorPresent)
                            //{
                            //    _writer.WriteStartElement("a", "buClr", OpenXmlNamespaces.DrawingML);
                            //    _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                            //    string s = defaultStyle.PRuns[tp.indentLevel].BulletColor.Red.ToString("X").PadLeft(2, '0') + defaultStyle.PRuns[tp.indentLevel].BulletColor.Green.ToString("X").PadLeft(2, '0') + defaultStyle.PRuns[tp.indentLevel].BulletColor.Blue.ToString("X").PadLeft(2, '0');
                            //    _writer.WriteAttributeString("val", s);
                            //    _writer.WriteEndElement();
                            //    _writer.WriteEndElement(); //buClr
                            //}
                            if (defaultStyle.PRuns[tp.indentLevel].BulletSizePresent)
                            {
                                if (defaultStyle.PRuns[tp.indentLevel].BulletSize > 0)
                                {
                                    this._writer.WriteStartElement("a", "buSzPct", OpenXmlNamespaces.DrawingML);
                                    this._writer.WriteAttributeString("val", (defaultStyle.PRuns[tp.indentLevel].BulletSize * 1000).ToString());
                                    this._writer.WriteEndElement(); //buSzPct
                                }
                                else
                                {
                                    //TODO
                                }
                            }
                            if (defaultStyle.PRuns[tp.indentLevel].BulletFontPresent)
                            {
                                this._writer.WriteStartElement("a", "buFont", OpenXmlNamespaces.DrawingML);
                                var fonts = this._ctx.Ppt.DocumentRecord.FirstChildWithType<b2xtranslator.PptFileFormat.Environment>().FirstChildWithType<FontCollection>();
                                var entity = fonts.entities[(int)defaultStyle.PRuns[tp.indentLevel].BulletTypefaceIdx];
                                if (entity.TypeFace.IndexOf('\0') > 0)
                                {
                                    this._writer.WriteAttributeString("typeface", entity.TypeFace.Substring(0, entity.TypeFace.IndexOf('\0')));
                                }
                                else
                                {
                                    this._writer.WriteAttributeString("typeface", entity.TypeFace);
                                }
                                this._writer.WriteEndElement(); //buChar
                            }


                            if (this.parentShapeTreeMapping != null && this.parentShapeTreeMapping.ShapeStyleTextProp9Atom != null && this.parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs.Count > runCount && this.parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs[runCount].fBulletHasAutoNumber == 1 && this.parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs[runCount].bulletAutoNumberScheme == -1)
                            {
                                this._writer.WriteStartElement("a", "buAutoNum", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteAttributeString("type", "arabicPeriod");
                                this._writer.WriteEndElement();
                            }
                            else if (defaultStyle.PRuns[tp.indentLevel].BulletCharPresent)
                            {
                                this._writer.WriteStartElement("a", "buChar", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteAttributeString("char", defaultStyle.PRuns[tp.indentLevel].BulletChar.ToString());
                                this._writer.WriteEndElement(); //buChar
                            }
                            else if (!defaultStyle.PRuns[tp.indentLevel].BulletCharPresent)
                            {
                                this._writer.WriteStartElement("a", "buChar", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteAttributeString("char", "•");
                                this._writer.WriteEndElement(); //buChar
                            }

                        }
                    }
                }
            }
            else
            {
                if (p.IndentLevel > 0) this._writer.WriteAttributeString("lvl", p.IndentLevel.ToString());
                if (p.LeftMarginPresent)
                {
                    this._writer.WriteAttributeString("marL", Utils.MasterCoordToEMU((int)p.LeftMargin).ToString());
                    writtenLeftMargin = (int)p.LeftMargin;
                } 
                else if (ruler != null && ruler.fLeftMargin1 && p.IndentLevel == 0){
                    this._writer.WriteAttributeString("marL", Utils.MasterCoordToEMU(ruler.leftMargin1).ToString());
                    writtenLeftMargin = ruler.leftMargin1;
                    if (!(p.IndentPresent || (defaultStyle != null && defaultStyle.PRuns.Count > p.IndentLevel && defaultStyle.PRuns[p.IndentLevel].IndentPresent) || (ruler != null && ruler.fIndent1 && p.IndentLevel == 0)))
                    {
                        this._writer.WriteAttributeString("indent", Utils.MasterCoordToEMU(-1 * ruler.leftMargin1).ToString());
                    }
                }
                else if (ruler != null && ruler.fLeftMargin2 && p.IndentLevel == 1)
                {
                    this._writer.WriteAttributeString("marL", Utils.MasterCoordToEMU(ruler.leftMargin2).ToString());
                    writtenLeftMargin = ruler.leftMargin2;
                    if (!(p.IndentPresent || (defaultStyle != null && defaultStyle.PRuns.Count > p.IndentLevel && defaultStyle.PRuns[p.IndentLevel].IndentPresent) || (ruler != null && ruler.fIndent2 && p.IndentLevel == 1)))
                    {
                        this._writer.WriteAttributeString("indent", Utils.MasterCoordToEMU(-1 * ruler.leftMargin1).ToString());
                    }
                }
                else if (ruler != null && ruler.fLeftMargin3 && p.IndentLevel == 2)
                {
                    this._writer.WriteAttributeString("marL", Utils.MasterCoordToEMU(ruler.leftMargin3).ToString());
                    writtenLeftMargin = ruler.leftMargin3;
                    if (!(p.IndentPresent || (defaultStyle != null && defaultStyle.PRuns.Count > p.IndentLevel && defaultStyle.PRuns[p.IndentLevel].IndentPresent) || (ruler != null && ruler.fIndent3 && p.IndentLevel == 2)))
                    {
                        this._writer.WriteAttributeString("indent", Utils.MasterCoordToEMU(-1 * ruler.leftMargin1).ToString());
                    }
                }
                else if (ruler != null && ruler.fLeftMargin4 && p.IndentLevel == 3)
                {
                    this._writer.WriteAttributeString("marL", Utils.MasterCoordToEMU(ruler.leftMargin4).ToString());
                    writtenLeftMargin = ruler.leftMargin4;
                    if (!(p.IndentPresent || (defaultStyle != null && defaultStyle.PRuns.Count > p.IndentLevel && defaultStyle.PRuns[p.IndentLevel].IndentPresent) || (ruler != null && ruler.fIndent4 && p.IndentLevel == 3)))
                    {
                        this._writer.WriteAttributeString("indent", Utils.MasterCoordToEMU(-1 * ruler.leftMargin1).ToString());
                    }
                }
                else if (ruler != null && ruler.fLeftMargin5 && p.IndentLevel == 4)
                {
                    this._writer.WriteAttributeString("marL", Utils.MasterCoordToEMU(ruler.leftMargin5).ToString());
                    writtenLeftMargin = ruler.leftMargin5;
                    if (!(p.IndentPresent || (defaultStyle != null && defaultStyle.PRuns.Count > p.IndentLevel && defaultStyle.PRuns[p.IndentLevel].IndentPresent) || (ruler != null && ruler.fIndent5 && p.IndentLevel == 4)))
                    {
                        this._writer.WriteAttributeString("indent", Utils.MasterCoordToEMU(-1 * ruler.leftMargin1).ToString());
                    }
                }
                else if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.dxTextLeft) && so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.TextBooleanProperties))
                {
                    var props = new TextBooleanProperties(so.OptionsByID[ShapeOptions.PropertyId.TextBooleanProperties].op);
                    if (props.fUsefAutoTextMargin && (props.fAutoTextMargin == false))
                    if (so.OptionsByID[ShapeOptions.PropertyId.dxTextLeft].op > 0)
                            this._writer.WriteAttributeString("marL", so.OptionsByID[ShapeOptions.PropertyId.dxTextLeft].op.ToString());
                    //writtenLeftMargin = Utils.EMUToMasterCoord((int)so.OptionsByID[ShapeOptions.PropertyId.dxTextLeft].op);
                }
                

                if (p.IndentPresent)
                {
                    this._writer.WriteAttributeString("indent", (-1 * (Utils.MasterCoordToEMU((int)(p.LeftMargin - p.Indent)))).ToString());
                }
                else if (ruler != null && ruler.fIndent1 && p.IndentLevel == 0)
                {
                    this._writer.WriteAttributeString("indent", (-1 * (Utils.MasterCoordToEMU((int)(ruler.leftMargin1 - ruler.indent1)))).ToString());
                }
                else if (ruler != null && ruler.fIndent2 && p.IndentLevel == 1)
                {
                    this._writer.WriteAttributeString("indent", (-1 * (Utils.MasterCoordToEMU((int)(ruler.leftMargin2 - ruler.indent2)))).ToString());
                }
                else if (ruler != null && ruler.fIndent3 && p.IndentLevel == 2)
                {
                    this._writer.WriteAttributeString("indent", (-1 * (Utils.MasterCoordToEMU((int)(ruler.leftMargin3 - ruler.indent3)))).ToString());
                }
                else if (ruler != null && ruler.fIndent4 && p.IndentLevel == 3)
                {
                    this._writer.WriteAttributeString("indent", (-1 * (Utils.MasterCoordToEMU((int)(ruler.leftMargin4 - ruler.indent4)))).ToString());
                }
                else if (ruler != null && ruler.fIndent5 && p.IndentLevel == 4)
                {
                    this._writer.WriteAttributeString("indent", (-1 * (Utils.MasterCoordToEMU((int)(ruler.leftMargin5 - ruler.indent5)))).ToString());
                }
                else if (defaultStyle != null && defaultStyle.PRuns.Count > p.IndentLevel && defaultStyle.PRuns[p.IndentLevel].IndentPresent)
                {
                    if (writtenLeftMargin == -1 )
                    {
                        writtenLeftMargin = (int)(defaultStyle.PRuns[p.IndentLevel].LeftMargin);
                    }
                    this._writer.WriteAttributeString("indent", (-1 * (Utils.MasterCoordToEMU((int)(writtenLeftMargin - defaultStyle.PRuns[p.IndentLevel].Indent)))).ToString());
                }
              
                if (p.AlignmentPresent)
                {
                    switch (p.Alignment)
                    {
                        case 0x0000: //Left
                            this._writer.WriteAttributeString("algn", "l");
                            break;
                        case 0x0001: //Center
                            this._writer.WriteAttributeString("algn", "ctr");
                            break;
                        case 0x0002: //Right
                            this._writer.WriteAttributeString("algn", "r");
                            break;
                        case 0x0003: //Justify
                            this._writer.WriteAttributeString("algn", "just");
                            break;
                        case 0x0004: //Distributed
                            this._writer.WriteAttributeString("algn", "dist");
                            break;
                        case 0x0005: //ThaiDistributed
                            this._writer.WriteAttributeString("algn", "thaiDist");
                            break;
                        case 0x0006: //JustifyLow
                            this._writer.WriteAttributeString("algn", "justLow");
                            break;
                    }
                }
                else if (defaultStyle != null && defaultStyle.PRuns.Count > p.IndentLevel)
                {
                    if (defaultStyle.PRuns[p.IndentLevel].AlignmentPresent)
                    {
                        switch (defaultStyle.PRuns[p.IndentLevel].Alignment)
                        {
                            case 0x0000: //Left
                                this._writer.WriteAttributeString("algn", "l");
                                break;
                            case 0x0001: //Center
                                this._writer.WriteAttributeString("algn", "ctr");
                                break;
                            case 0x0002: //Right
                                this._writer.WriteAttributeString("algn", "r");
                                break;
                            case 0x0003: //Justify
                                this._writer.WriteAttributeString("algn", "just");
                                break;
                            case 0x0004: //Distributed
                                this._writer.WriteAttributeString("algn", "dist");
                                break;
                            case 0x0005: //ThaiDistributed
                                this._writer.WriteAttributeString("algn", "thaiDist");
                                break;
                            case 0x0006: //JustifyLow
                                this._writer.WriteAttributeString("algn", "justLow");
                                break;
                        }
                    }
                }

                if (p.DefaultTabSizePresent)
                {
                    this._writer.WriteAttributeString("defTabSz", Utils.MasterCoordToEMU((int)p.DefaultTabSize).ToString());
                }
                else if (defaultStyle != null && defaultStyle.PRuns.Count > p.IndentLevel)
                {
                    if (defaultStyle.PRuns[p.IndentLevel].DefaultTabSizePresent)
                    {
                        this._writer.WriteAttributeString("defTabSz", Utils.MasterCoordToEMU((int)defaultStyle.PRuns[p.IndentLevel].DefaultTabSize).ToString());
                    }
                }

                if (p.LineSpacingPresent)
                {
                    this._writer.WriteStartElement("a", "lnSpc", OpenXmlNamespaces.DrawingML);
                    //_writer.WriteStartElement("a", "spcPct", OpenXmlNamespaces.DrawingML);
                    //_writer.WriteAttributeString("val", (p.LineSpacing * 1000).ToString());
                    //_writer.WriteEndElement(); //spcPct

                    if (p.LineSpacing < 0)
                    {
                        this._writer.WriteStartElement("a", "spcPts", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("val", (-1 * p.LineSpacing * 12).ToString()); //TODO: this has to be verified!
                        this._writer.WriteEndElement(); //spcPct
                    }
                    else
                    {
                        this._writer.WriteStartElement("a", "spcPct", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("val", (1000 * p.LineSpacing).ToString());
                        this._writer.WriteEndElement(); //spcPct
                    }


                    this._writer.WriteEndElement(); //lnSpc
                }
                if (p.SpaceBeforePresent)
                {
                    this._writer.WriteStartElement("a", "spcBef", OpenXmlNamespaces.DrawingML);
                    if (p.SpaceBefore < 0)
                    {
                        this._writer.WriteStartElement("a", "spcPts", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("val", (-1 * 12 * p.SpaceBefore).ToString()); //TODO: the 12 is wrong: find correct value
                        this._writer.WriteEndElement(); //spcPct
                    }
                    else
                    {
                        this._writer.WriteStartElement("a", "spcPct", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("val", (1000 * p.SpaceBefore).ToString());
                        this._writer.WriteEndElement(); //spcPct
                    }
                    this._writer.WriteEndElement(); //spcBef
                }

                if (p.SpaceAfterPresent)
                {
                    this._writer.WriteStartElement("a", "spcAft", OpenXmlNamespaces.DrawingML);
                    if (p.SpaceAfter < 0)
                    {
                        this._writer.WriteStartElement("a", "spcPts", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("val", (-1 * 12 * p.SpaceAfter).ToString()); //TODO: the 12 is wrong: find correct value
                        this._writer.WriteEndElement(); //spcPct
                    }
                    else
                    {
                        this._writer.WriteStartElement("a", "spcPct", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("val", (1000 * p.SpaceAfter).ToString());
                        this._writer.WriteEndElement(); //spcPct
                    }
                    this._writer.WriteEndElement(); //spcAft
                }

                if (p.BulletFlagsFieldPresent)
                {
                    if ((p.BulletFlags & (ushort)ParagraphMask.HasBullet) == 0)
                    {
                        this._writer.WriteElementString("a", "buNone", OpenXmlNamespaces.DrawingML, "");
                    }
                    else
                    {
                                                
                        if (p.BulletColorPresent)
                        {
                            this._writer.WriteStartElement("a", "buClr", OpenXmlNamespaces.DrawingML);
                            
                            string s = p.BulletColor.Red.ToString("X").PadLeft(2, '0') + p.BulletColor.Green.ToString("X").PadLeft(2, '0') + p.BulletColor.Blue.ToString("X").PadLeft(2, '0');
                            switch (p.BulletColor.Index)
                            {
                                case 7:
                                    this._writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);
                                    this._writer.WriteAttributeString("val", "folHlink");
                                    this._writer.WriteEndElement();
                                    break;
                                case 6:
                                    this._writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);
                                    this._writer.WriteAttributeString("val", "hlink");
                                    this._writer.WriteEndElement();
                                    break;
                                case 3:
                                    this._writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);
                                    this._writer.WriteAttributeString("val", "tx2");
                                    this._writer.WriteEndElement();
                                    break;
                                case 2:
                                    this._writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);
                                    this._writer.WriteAttributeString("val", "bg2");
                                    this._writer.WriteEndElement();
                                    break;
                                case 1:
                                    this._writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);
                                    this._writer.WriteAttributeString("val", "bg1");
                                    this._writer.WriteEndElement();
                                    break;
                                default:
                                    this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                                    this._writer.WriteAttributeString("val", s);
                                    this._writer.WriteEndElement();
                                    break;
                            }

                            this._writer.WriteEndElement(); //buClr
                        }
                        if (p.BulletSizePresent)
                        {
                            if (p.BulletSize > 0)
                            {
                                this._writer.WriteStartElement("a", "buSzPct", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteAttributeString("val", (p.BulletSize * 1000).ToString());
                                this._writer.WriteEndElement(); //buSzPct
                            }
                            else
                            {
                                //TODO
                            }
                        }
                        else if (p.BulletFlagsFieldPresent && (p.BulletFlags & 0x1 << 3) > 0)
                        {
                            this._writer.WriteStartElement("a", "buSzPct", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("val", "75000");
                            this._writer.WriteEndElement(); //buSzPct
                        }

                        if (p.BulletFontPresent)
                        {
                            if (!(p.BulletFlagsFieldPresent && (p.BulletFlags & 0x1 << 1) == 0))
                            {
                                this._writer.WriteStartElement("a", "buFont", OpenXmlNamespaces.DrawingML);
                                var fonts = this._ctx.Ppt.DocumentRecord.FirstChildWithType<b2xtranslator.PptFileFormat.Environment>().FirstChildWithType<FontCollection>();
                                var entity = fonts.entities[(int)p.BulletTypefaceIdx];
                                if (entity.TypeFace.IndexOf('\0') > 0)
                                {
                                    this._writer.WriteAttributeString("typeface", entity.TypeFace.Substring(0, entity.TypeFace.IndexOf('\0')));
                                }
                                else
                                {
                                    this._writer.WriteAttributeString("typeface", entity.TypeFace);
                                }
                                this._writer.WriteEndElement(); //buChar
                            }
                            else
                            {
                                this._writer.WriteElementString("a", "buFontTx", OpenXmlNamespaces.DrawingML, "");
                            }
                        }

                        bool autoNumberingWritten = false;
                        bool bulletWritten = false;
                        if (this._ctx.Ppt.DocumentRecord.DocInfoListContainer.FirstDescendantWithType<OutlineTextProps9Container>() != null)
                        {
                            var c = this._ctx.Ppt.DocumentRecord.DocInfoListContainer.FirstDescendantWithType<OutlineTextProps9Container>();
                            var slide = so.FirstAncestorWithType<Slide>();

                            if (slide != null)
                            foreach (var entry in c.OutlineTextProps9Entries)
                            {
                                if (slide.PersistAtom.SlideId == entry.outlineTextHeaderAtom.slideIdRef)
                                {
                                    if (entry.styleTextProp9Atom.P9Runs.Count > runCount && entry.styleTextProp9Atom.P9Runs[runCount].fBulletHasAutoNumber == 1)
                                    {
                                        switch (entry.styleTextProp9Atom.P9Runs[runCount].bulletAutoNumberScheme)
                                        {
                                            case -1:
                                            case 3:
                                                    this._writer.WriteStartElement("a", "buAutoNum", OpenXmlNamespaces.DrawingML);
                                                    this._writer.WriteAttributeString("type", "arabicPeriod");
                                                if (entry.styleTextProp9Atom.P9Runs[runCount].startAt != -1)
                                                {
                                                        this._writer.WriteAttributeString("startAt", entry.styleTextProp9Atom.P9Runs[runCount].startAt.ToString());
                                                }
                                                    this._writer.WriteEndElement();
                                                autoNumberingWritten = true;
                                                break;
                                            case 1:
                                                    this._writer.WriteStartElement("a", "buAutoNum", OpenXmlNamespaces.DrawingML);
                                                    this._writer.WriteAttributeString("type", "alphaUcPeriod");
                                                if (entry.styleTextProp9Atom.P9Runs[runCount].startAt != -1)
                                                {
                                                        this._writer.WriteAttributeString("startAt", entry.styleTextProp9Atom.P9Runs[runCount].startAt.ToString());
                                                }
                                                    this._writer.WriteEndElement();
                                                autoNumberingWritten = true;
                                                break;
                                        }
                                    }
                                    else if (entry.styleTextProp9Atom.P9Runs.Count > runCount && entry.styleTextProp9Atom.P9Runs[runCount].BulletBlipReferencePresent)
                                    {
                                       var blips = ((RegularContainer)c.ParentRecord).FirstChildWithType<BlipCollection9Container>();
                                        if (blips != null && blips.Children.Count > 0)
                                        {
                                            ImagePart imgPart = null;

                                            var b = ((BlipEntityAtom)blips.Children[entry.styleTextProp9Atom.P9Runs[runCount].bulletblipref]).blip;

                                            if (b == null)
                                            {
                                                var mb = ((BlipEntityAtom)blips.Children[0]).mblip;
                                                imgPart = this.parentShapeTreeMapping.parentSlideMapping.targetPart.AddImagePart(ShapeTreeMapping.getImageType(mb.TypeCode));
                                                imgPart.TargetDirectory = "..\\media";
                                                var outStream = imgPart.GetStream();
                                                var decompressed = mb.Decrompress();
                                                outStream.Write(decompressed, 0, decompressed.Length);
                                            }
                                            else
                                            {
                                                imgPart = this.parentShapeTreeMapping.parentSlideMapping.targetPart.AddImagePart(ShapeTreeMapping.getImageType(b.TypeCode));
                                                imgPart.TargetDirectory = "..\\media";
                                                var outStream = imgPart.GetStream();
                                                outStream.Write(b.m_pvBits, 0, b.m_pvBits.Length);
                                            }

                                                this._writer.WriteStartElement("a", "buBlip", OpenXmlNamespaces.DrawingML);
                                                this._writer.WriteStartElement("a", "blip", OpenXmlNamespaces.DrawingML);
                                                this._writer.WriteAttributeString("r", "embed", OpenXmlNamespaces.Relationships, imgPart.RelIdToString);
                                                this._writer.WriteEndElement(); //blip
                                                this._writer.WriteEndElement(); //buBlip
                                            bulletWritten = true;
                                        }
                                    }
                                }
                            }

                            //OutlineTextPropsHeader9Atom a = c.FirstChildWithType<OutlineTextPropsHeader9Atom>();
                            //Slide slide = so.FirstAncestorWithType<Slide>();
                            //if (slide.PersistAtom.SlideId == a.slideIdRef)
                            //{
                            //    StyleTextProp9Atom s = c.FirstChildWithType<StyleTextProp9Atom>();
                            //    if (s.P9Runs.Count > runCount && s.P9Runs[runCount].fBulletHasAutoNumber == 1)
                            //    {
                            //        switch (s.P9Runs[runCount].bulletAutoNumberScheme)
                            //        {
                            //            case -1:
                            //                _writer.WriteStartElement("a", "buAutoNum", OpenXmlNamespaces.DrawingML);
                            //                _writer.WriteAttributeString("type", "arabicPeriod");
                            //                _writer.WriteEndElement();
                            //                autoNumberingWritten = true;
                            //                break;
                            //            case 1:
                            //                _writer.WriteStartElement("a", "buAutoNum", OpenXmlNamespaces.DrawingML);
                            //                _writer.WriteAttributeString("type", "alphaUcPeriod");
                            //                _writer.WriteEndElement();
                            //                autoNumberingWritten = true;
                            //                break;
                            //        }
                            //    }
                            //}
                        }

                        if (!autoNumberingWritten)
                        {
                            if (this.parentShapeTreeMapping != null && this.parentShapeTreeMapping.ShapeStyleTextProp9Atom != null && this.parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs.Count > runCount && this.parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs[runCount].fBulletHasAutoNumber == 1)
                            {
                                switch(this.parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs[runCount].bulletAutoNumberScheme)
                                {
                                    case -1:
                                    case 3:
                                        this._writer.WriteStartElement("a", "buAutoNum", OpenXmlNamespaces.DrawingML);
                                        this._writer.WriteAttributeString("type", "arabicPeriod");
                                        if (this.parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs[runCount].startAt != -1)
                                        {
                                            this._writer.WriteAttributeString("startAt", this.parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs[runCount].startAt.ToString());
                                        }
                                        this._writer.WriteEndElement();
                                        break;
                                    case 1:
                                        this._writer.WriteStartElement("a", "buAutoNum", OpenXmlNamespaces.DrawingML);
                                        this._writer.WriteAttributeString("type", "alphaUcPeriod");
                                        if (this.parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs[runCount].startAt != -1)
                                        {
                                            this._writer.WriteAttributeString("startAt", this.parentShapeTreeMapping.ShapeStyleTextProp9Atom.P9Runs[runCount].startAt.ToString());
                                        }
                                        this._writer.WriteEndElement();
                                        break;
                                }
                            }
                            else if (!bulletWritten && p.BulletCharPresent)
                            {
                                this._writer.WriteStartElement("a", "buChar", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteAttributeString("char", p.BulletChar.ToString());
                                this._writer.WriteEndElement(); //buChar

                                var s = so.FirstAncestorWithType<Slide>();
                            }
                            else if (!bulletWritten && !p.BulletCharPresent)
                            {
                                this._writer.WriteStartElement("a", "buChar", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteAttributeString("char", "•");
                                this._writer.WriteEndElement(); //buChar
                            }
                        }
                    }
                }
            }

            this._writer.WriteEndElement(); //pPr
        }
                
    }
}
