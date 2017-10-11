

using System.Collections.Generic;
using b2xtranslator.PptFileFormat;
using b2xtranslator.CommonTranslatorLib;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.OfficeDrawing;
using b2xtranslator.Tools;

namespace b2xtranslator.PresentationMLMapping
{
    class TextMasterStyleMapping :
        AbstractOpenXmlMapping
    {
        protected ConversionContext _ctx;
        public RegularContainer _Master;
        public PresentationMapping<RegularContainer> _parentSlideMapping = null;

        private int lastSpaceBefore = 0;
        private string lastColor = "";
        private string lastBulletFont = "";
        private string lastBulletChar = "";
        private string lastBulletColor = "";
        private string lastSize = "";
        private string lastTypeface = "";
        public TextMasterStyleMapping(ConversionContext ctx, XmlWriter writer, PresentationMapping<RegularContainer> parentSlideMapping)
            : base(writer)
        {
            this._ctx = ctx;
            this._parentSlideMapping = parentSlideMapping;
        }

        public List<TextMasterStyleAtom> titleAtoms = new List<TextMasterStyleAtom>();
        public List<TextMasterStyleAtom> bodyAtoms = new List<TextMasterStyleAtom>();
        public List<TextMasterStyleAtom> CenterBodyAtoms = new List<TextMasterStyleAtom>();
        public List<TextMasterStyleAtom> CenterTitleAtoms = new List<TextMasterStyleAtom>();
        public List<TextMasterStyleAtom> noteAtoms = new List<TextMasterStyleAtom>();
        public void Apply(RegularContainer Master)
        {
            this._Master = Master;

            var atoms = Master.AllChildrenWithType<TextMasterStyleAtom>();            

            var body9atoms = new List<TextMasterStyle9Atom>();
            var title9atoms = new List<TextMasterStyle9Atom>();
            foreach (var progtags in Master.AllChildrenWithType<ProgTags>())
	        {
        		foreach (var progbinarytag in progtags.AllChildrenWithType<ProgBinaryTag>())
	            {
                    foreach (var blob in progbinarytag.AllChildrenWithType<ProgBinaryTagDataBlob>())
                    {
                        foreach (var atom in blob.AllChildrenWithType<TextMasterStyle9Atom>())
                        {
                            if (atom.Instance == 0) title9atoms.Add(atom);
                            if (atom.Instance == 1) body9atoms.Add(atom);
                        }
                    }            		
	            }
	        }
            
            foreach (var atom in atoms)
            {
                if (atom.Instance == 0) this.titleAtoms.Add(atom);   
                if (atom.Instance == 1) this.bodyAtoms.Add(atom);
                if (atom.Instance == 5) this.CenterBodyAtoms.Add(atom);
                if (atom.Instance == 6) this.CenterTitleAtoms.Add(atom);
            }

            this._writer.WriteStartElement("p", "txStyles", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "titleStyle", OpenXmlNamespaces.PresentationML);

            ParagraphRun9 pr9 = null;
            foreach (var atom in this.titleAtoms)
            {
                this.lastSpaceBefore = 0;
                this.lastBulletFont = "";
                this.lastBulletChar = "";
                this.lastColor = "";
                this.lastBulletColor = "";
                this.lastSize = "";
                for (int i = 0; i < atom.IndentLevelCount; i++)
                {
                    pr9 = null;
                    if (title9atoms.Count > 0 && title9atoms[0].pruns.Count > i) pr9 = title9atoms[0].pruns[i];
                    writepPr(atom.CRuns[i], atom.PRuns[i], pr9, i, true);
                }
                for (int i = atom.IndentLevelCount; i < 9; i++)
                {
                    pr9 = null;
                    if (title9atoms.Count > 0 && title9atoms[0].pruns.Count > i) pr9 = title9atoms[0].pruns[i];
                    writepPr(atom.CRuns[0], atom.PRuns[0], pr9, i, true);
                }
            }

            this._writer.WriteEndElement(); //titleStyle

            this._writer.WriteStartElement("p", "bodyStyle", OpenXmlNamespaces.PresentationML);

            foreach (var atom in this.bodyAtoms)
            {
                this.lastSpaceBefore = 0;
                this.lastColor = "";
                this.lastBulletFont = "";
                this.lastBulletChar = "";
                this.lastBulletColor = "";
                this.lastSize = "";
                for (int i = 0; i < atom.IndentLevelCount; i++)
                {
                    pr9 = null;
                    if (body9atoms.Count > 0 && body9atoms[0].pruns.Count > i) pr9 = body9atoms[0].pruns[i];
                    writepPr(atom.CRuns[i], atom.PRuns[i], pr9, i, false);
                }
                for (int i = atom.IndentLevelCount; i < 9; i++)
                {
                    pr9 = null;
                    if (body9atoms.Count > 0 && body9atoms[0].pruns.Count > i) pr9 = body9atoms[0].pruns[i];
                    writepPr(atom.CRuns[0], atom.PRuns[0],pr9, i, false);
                }
            }

            this._writer.WriteEndElement(); //bodyStyle

            this._writer.WriteEndElement(); //txStyles
        }

        public void ApplyNotesMaster(RegularContainer notesMaster)
        {
            this._Master = notesMaster;
            var m = this._ctx.Ppt.MainMasterRecords[0];
            var atoms = m.AllChildrenWithType<TextMasterStyleAtom>();
            foreach (var atom in atoms)
            {
                if (atom.Instance == 2) this.noteAtoms.Add(atom);
            }

            this._writer.WriteStartElement("p", "notesStyle", OpenXmlNamespaces.PresentationML);

            ParagraphRun9 pr9 = null;
            foreach (var atom in this.noteAtoms)
            {
                this.lastSpaceBefore = 0;
                this.lastBulletFont = "";
                this.lastBulletChar = "";
                this.lastBulletColor = "";
                this.lastColor = "";
                this.lastSize = "";
                for (int i = 0; i < atom.IndentLevelCount; i++)
                {
                    pr9 = null;
                    writepPr(atom.CRuns[i], atom.PRuns[i], pr9, i, true);
                }
                for (int i = atom.IndentLevelCount; i < 9; i++)
                {
                    pr9 = null;
                    writepPr(atom.CRuns[0], atom.PRuns[0], pr9, i, true);
                }
            }

            this._writer.WriteEndElement();
        }

        public void ApplyHandoutMaster(RegularContainer handoutMaster)
        {
            this._Master = handoutMaster;
            var m = this._ctx.Ppt.MainMasterRecords[0];
            var atoms = m.AllChildrenWithType<TextMasterStyleAtom>();
            foreach (var atom in atoms)
            {
                if (atom.Instance == 2) this.noteAtoms.Add(atom);
            }

            this._writer.WriteStartElement("p", "handoutStyle", OpenXmlNamespaces.PresentationML);

            ParagraphRun9 pr9 = null;
            foreach (var atom in this.noteAtoms)
            {
                this.lastSpaceBefore = 0;
                this.lastBulletFont = "";
                this.lastBulletChar = "";
                this.lastBulletColor = "";
                this.lastColor = "";
                this.lastSize = "";
                for (int i = 0; i < atom.IndentLevelCount; i++)
                {
                    pr9 = null;
                    writepPr(atom.CRuns[i], atom.PRuns[i], pr9, i, true);
                }
                for (int i = atom.IndentLevelCount; i < 9; i++)
                {
                    pr9 = null;
                    writepPr(atom.CRuns[0], atom.PRuns[0], pr9, i, true);
                }
            }

            this._writer.WriteEndElement();
        }

        public void writepPr(CharacterRun cr, ParagraphRun pr, ParagraphRun9 pr9, int IndentLevel, bool isTitle)
        {
            writepPr(cr, pr, pr9, IndentLevel, isTitle, false);
        }

        public void writepPr(CharacterRun cr, ParagraphRun pr, ParagraphRun9 pr9, int IndentLevel, bool isTitle, bool isDefault)
        {

            //TextMasterStyleAtom defaultStyle = _ctx.Ppt.DocumentRecord.FirstChildWithType<b2xtranslator.PptFileFormat.Environment>().FirstChildWithType<TextMasterStyleAtom>();

            this._writer.WriteStartElement("a", "lvl" + (IndentLevel+1).ToString() + "pPr", OpenXmlNamespaces.DrawingML);

            //marL
            if (pr.LeftMarginPresent && !isDefault) this._writer.WriteAttributeString("marL", Utils.MasterCoordToEMU((int)pr.LeftMargin).ToString());
            //marR
            //lvl
            if (pr.IndentLevel > 0) this._writer.WriteAttributeString("lvl", pr.IndentLevel.ToString());
            //indent
            if (pr.IndentPresent && !isDefault) this._writer.WriteAttributeString("indent", (-1 * (Utils.MasterCoordToEMU((int)(pr.LeftMargin - pr.Indent)))).ToString());
            //algn
            if (pr.AlignmentPresent)
            {
                switch (pr.Alignment)
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
            //defTabSz
            if (pr.DefaultTabSizePresent)
            {
                this._writer.WriteAttributeString("defTabSz", Utils.MasterCoordToEMU((int)pr.DefaultTabSize).ToString());
            }
            //rtl
            if (pr.TextDirectionPresent)
            {
                switch (pr.TextDirection)
                {
                    case 0x0000:
                        this._writer.WriteAttributeString("rtl", "0");
                        break;
                    case 0x0001:
                        this._writer.WriteAttributeString("rtl", "1");
                        break;
                }
            }
            else
            {
                this._writer.WriteAttributeString("rtl", "0");
            }
            //eaLnkBrk
            //fontAlgn
            if (pr.FontAlignPresent)
            {
                switch (pr.FontAlign)
                {
                    case 0x0000: //Roman
                        this._writer.WriteAttributeString("fontAlgn", "base");
                        break;
                    case 0x0001: //Hanging
                        this._writer.WriteAttributeString("fontAlgn", "t");
                        break;
                    case 0x0002: //Center
                        this._writer.WriteAttributeString("fontAlgn", "ctr");
                        break;
                    case 0x0003: //UpholdFixed
                        this._writer.WriteAttributeString("fontAlgn", "b");
                        break;
                }
            }
            //latinLnBrk
            //hangingPunct


            //lnSpc
            //spcBef
            if (pr.SpaceBeforePresent)
            {
                this._writer.WriteStartElement("a", "spcBef", OpenXmlNamespaces.DrawingML);
                if (pr.SpaceBefore < 0)
                {
                    this._writer.WriteStartElement("a", "spcPts", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("val", (-1 * 12 * pr.SpaceBefore).ToString()); //TODO: the 12 is wrong
                    this._writer.WriteEndElement(); //spcPct
                }
                else
                {
                    this._writer.WriteStartElement("a", "spcPct", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("val", (1000 * pr.SpaceBefore).ToString());
                    this._writer.WriteEndElement(); //spcPct
                }
                this._writer.WriteEndElement(); //spcBef
                this.lastSpaceBefore = (int)pr.SpaceBefore;
            }
            else
            {
                if (this.lastSpaceBefore != 0)
                {
                    this._writer.WriteStartElement("a", "spcBef", OpenXmlNamespaces.DrawingML);
                    if (this.lastSpaceBefore < 0)
                    {
                        this._writer.WriteStartElement("a", "spcPts", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("val", (-1 * 12 * this.lastSpaceBefore).ToString()); //TODO: the 12 is wrong
                        this._writer.WriteEndElement(); //spcPct
                    }
                    else
                    {
                        this._writer.WriteStartElement("a", "spcPct", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("val", (1000 * this.lastSpaceBefore).ToString());
                        this._writer.WriteEndElement(); //spcPct
                    }
                    this._writer.WriteEndElement(); //spcBef
                }
            }
            //spcAft
            if (pr.SpaceAfterPresent)
            {
                this._writer.WriteStartElement("a", "spcAft", OpenXmlNamespaces.DrawingML);
                if (pr.SpaceAfter < 0)
                {
                    this._writer.WriteStartElement("a", "spcPts", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("val", (-1 * pr.SpaceAfter).ToString()); //TODO: this has to be verified!
                    this._writer.WriteEndElement(); //spcPct
                }
                else
                {
                    this._writer.WriteStartElement("a", "spcPct", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("val", pr.SpaceAfter.ToString());
                    this._writer.WriteEndElement(); //spcPct
                }
                this._writer.WriteEndElement(); //spcAft
            }
            //EG_TextBulletColor
            //EG_TextBulletSize
            //EG_TextBulletTypeFace
            //EG_TextBullet


            bool bulletwritten = false;
            if (pr9 != null)
            {
                if (pr9.BulletBlipReferencePresent)
                    foreach (var progtags in this._ctx.Ppt.DocumentRecord.FirstChildWithType<List>().AllChildrenWithType<ProgTags>())
                    {
                        foreach (var bintags in progtags.AllChildrenWithType<ProgBinaryTag>())
                        {
                            foreach (var data in bintags.AllChildrenWithType<ProgBinaryTagDataBlob>())
                            {
                                foreach (var blips in data.AllChildrenWithType<BlipCollection9Container>())
                                {
                                    if (blips.Children.Count > pr9.bulletblipref & pr9.bulletblipref > -1)
                                    {
                                        ImagePart imgPart = null;

                                        var b = ((BlipEntityAtom)blips.Children[pr9.bulletblipref]).blip;

                                        if (b == null)
                                        {
                                            var mb = ((BlipEntityAtom)blips.Children[pr9.bulletblipref]).mblip;
                                            imgPart = this._parentSlideMapping.targetPart.AddImagePart(ShapeTreeMapping.getImageType(mb.TypeCode));
                                            imgPart.TargetDirectory = "..\\media";
                                            var outStream = imgPart.GetStream();
                                            var decompressed = mb.Decrompress();
                                            outStream.Write(decompressed, 0, decompressed.Length);
                                            //outStream.Write(mb.m_pvBits, 0, mb.m_pvBits.Length);
                                        }
                                        else
                                        {
                                            imgPart = this._parentSlideMapping.targetPart.AddImagePart(ShapeTreeMapping.getImageType(b.TypeCode));
                                            imgPart.TargetDirectory = "..\\media";
                                            var outStream = imgPart.GetStream();
                                            outStream.Write(b.m_pvBits, 0, b.m_pvBits.Length);
                                        }

                                        this._writer.WriteStartElement("a", "buBlip", OpenXmlNamespaces.DrawingML);
                                        this._writer.WriteStartElement("a", "blip", OpenXmlNamespaces.DrawingML);
                                        this._writer.WriteAttributeString("r", "embed", OpenXmlNamespaces.Relationships, imgPart.RelIdToString);
                                        this._writer.WriteEndElement(); //blip
                                        this._writer.WriteEndElement(); //buBlip
                                        bulletwritten = true;
                                    }
                                }
                            }
                        }
                    }
            }

            if (!bulletwritten & !isTitle)
            {
                if (pr.BulletFlagsFieldPresent & (pr.BulletFlags & (ushort)ParagraphMask.HasBullet) == 0)
                {
                    this._writer.WriteElementString("a", "buNone", OpenXmlNamespaces.DrawingML, "");
                }
                else
                {
                    if (pr.BulletColorPresent && (!(pr.BulletFlagsFieldPresent && (pr.BulletFlags & 1 << 2) == 0)))
                    {
                        writeBuClr((RegularContainer)this._Master, pr.BulletColor, ref this.lastBulletColor);
                    }
                    else if (this.lastBulletColor.Length > 0)
                    {
                        this._writer.WriteStartElement("a", "buClr", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("val", this.lastBulletColor);
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement(); //buClr
                    }

                    if (pr.BulletSizePresent)
                    {
                        if (pr.BulletSize > 0 && pr.BulletSize != 100)
                        {
                            this._writer.WriteStartElement("a", "buSzPct", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("val", (pr.BulletSize * 1000).ToString());
                            this._writer.WriteEndElement(); //buSzPct
                        }
                        else
                        {
                            //TODO
                        }
                     }



                     if (pr.BulletFontPresent)
                     {
                        this._writer.WriteStartElement("a", "buFont", OpenXmlNamespaces.DrawingML);
                        var fonts = this._ctx.Ppt.DocumentRecord.FirstChildWithType<b2xtranslator.PptFileFormat.Environment>().FirstChildWithType<FontCollection>();
                        var entity = fonts.entities[(int)pr.BulletTypefaceIdx];
                        if (entity.TypeFace.IndexOf('\0') > 0)
                        {
                            this._writer.WriteAttributeString("typeface", entity.TypeFace.Substring(0, entity.TypeFace.IndexOf('\0')));
                        }
                        else
                        {
                            this._writer.WriteAttributeString("typeface", entity.TypeFace);
                        }
                        this._writer.WriteEndElement(); //buChar
                        this.lastBulletFont = entity.TypeFace;
                     }
                     else if (this.lastBulletFont.Length > 0)
                     {
                        this._writer.WriteStartElement("a", "buFont", OpenXmlNamespaces.DrawingML);
                         if (this.lastBulletFont.IndexOf('\0') > 0)
                         {
                            this._writer.WriteAttributeString("typeface", this.lastBulletFont.Substring(0, this.lastBulletFont.IndexOf('\0')));
                         }
                         else
                         {
                            this._writer.WriteAttributeString("typeface", this.lastBulletFont);
                         }
                        this._writer.WriteEndElement(); //buChar
                     }
                     if (pr.BulletCharPresent)
                     {
                        this._writer.WriteStartElement("a", "buChar", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("char", pr.BulletChar.ToString());
                        this._writer.WriteEndElement(); //buChar
                        this.lastBulletChar = pr.BulletChar.ToString();
                     }
                     else if (this.lastBulletChar.Length > 0)
                     {
                        this._writer.WriteStartElement("a", "buChar", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("char", this.lastBulletChar);
                        this._writer.WriteEndElement(); //buChar
                     }
                 }
            }
            
            //tabLst
            //defRPr
            //extLst

            new CharacterRunPropsMapping(this._ctx, this._writer).Apply(cr, "defRPr", (RegularContainer)this._Master, ref this.lastColor, ref this.lastSize, ref this.lastTypeface, "", "", null,IndentLevel,null,null,0, false);

            this._writer.WriteEndElement(); //lvlXpPr
        }

        public void writeBuClr(RegularContainer slide, GrColorAtom color, ref string lastColor)
        {
            if (color.IsSchemeColor) //TODO: to be fully implemented
            {
                //_writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);

                if (slide == null)
                {
                    ////TODO: what shall be used in this case (happens for default text style in presentation.xml)
                    //_writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    //_writer.WriteAttributeString("val", "000000");
                    //_writer.WriteEndElement();

                    this._writer.WriteStartElement("a", "buClr", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);
                    switch (color.Index)
                    {
                        case 0x00:
                            this._writer.WriteAttributeString("val", "bg1"); //background
                            break;
                        case 0x01:
                            this._writer.WriteAttributeString("val", "tx1"); //text
                            break;
                        case 0x02:
                            this._writer.WriteAttributeString("val", "dk1"); //shadow
                            break;
                        case 0x03:
                            this._writer.WriteAttributeString("val", "tx1"); //title text
                            break;
                        case 0x04:
                            this._writer.WriteAttributeString("val", "bg2"); //fill
                            break;
                        case 0x05:
                            this._writer.WriteAttributeString("val", "accent1"); //accent1
                            break;
                        case 0x06:
                            this._writer.WriteAttributeString("val", "accent2"); //accent2
                            break;
                        case 0x07:
                            this._writer.WriteAttributeString("val", "accent3"); //accent3
                            break;
                        case 0xFE: //sRGB
                            lastColor = color.Red.ToString("X").PadLeft(2, '0') + color.Green.ToString("X").PadLeft(2, '0') + color.Blue.ToString("X").PadLeft(2, '0');
                            this._writer.WriteAttributeString("val", lastColor);
                            break;
                        case 0xFF: //undefined
                            break;
                    }
                    this._writer.WriteEndElement();
                    this._writer.WriteEndElement();

                }
                else
                {
                    ColorSchemeAtom MasterScheme = null;
                    var ato = slide.FirstChildWithType<SlideAtom>();
                    List<ColorSchemeAtom> colors;
                    if (ato != null && Tools.Utils.BitmaskToBool(ato.Flags, 0x1 << 1) && ato.MasterId != 0)
                    {
                        colors = this._ctx.Ppt.FindMasterRecordById(ato.MasterId).AllChildrenWithType<ColorSchemeAtom>();
                    }
                    else
                    {
                        colors = slide.AllChildrenWithType<ColorSchemeAtom>();
                    }
                    foreach (var colorsch in colors)
                    {
                        if (colorsch.Instance == 1) MasterScheme = colorsch;
                    }

                    if (color.Index != 0xFF)
                    {
                        this._writer.WriteStartElement("a", "buClr", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                        switch (color.Index)
                        {
                            case 0x00: //background
                                lastColor = new RGBColor(MasterScheme.Background, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                                this._writer.WriteAttributeString("val", lastColor);
                                break;
                            case 0x01: //text
                                lastColor = new RGBColor(MasterScheme.TextAndLines, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                                this._writer.WriteAttributeString("val", lastColor);
                                break;
                            case 0x02: //shadow
                                lastColor = new RGBColor(MasterScheme.Shadows, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                                this._writer.WriteAttributeString("val", lastColor);
                                break;
                            case 0x03: //title
                                lastColor = new RGBColor(MasterScheme.TitleText, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                                this._writer.WriteAttributeString("val", lastColor);
                                break;
                            case 0x04: //fill
                                lastColor = new RGBColor(MasterScheme.Fills, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                                this._writer.WriteAttributeString("val", lastColor);
                                break;
                            case 0x05: //accent1
                                lastColor = new RGBColor(MasterScheme.Accent, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                                this._writer.WriteAttributeString("val", lastColor);
                                break;
                            case 0x06: //accent2
                                lastColor = new RGBColor(MasterScheme.AccentAndHyperlink, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                                this._writer.WriteAttributeString("val", lastColor);
                                break;
                            case 0x07: //accent3
                                lastColor = new RGBColor(MasterScheme.AccentAndFollowedHyperlink, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                                this._writer.WriteAttributeString("val", lastColor);
                                break;
                            case 0xFE: //sRGB
                                lastColor = color.Red.ToString("X").PadLeft(2, '0') + color.Green.ToString("X").PadLeft(2, '0') + color.Blue.ToString("X").PadLeft(2, '0');
                                this._writer.WriteAttributeString("val", lastColor);
                                break;
                            case 0xFF: //undefined
                            default:
                                break;

                        }
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                    }
                    //_writer.WriteEndElement();
                }
            }
            else
            {
                this._writer.WriteStartElement("a", "buClr", OpenXmlNamespaces.DrawingML);
                this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                lastColor = color.Red.ToString("X").PadLeft(2, '0') + color.Green.ToString("X").PadLeft(2, '0') + color.Blue.ToString("X").PadLeft(2, '0');
                this._writer.WriteAttributeString("val", lastColor);
                this._writer.WriteEndElement();
                this._writer.WriteEndElement();
            }
            
        }

    }
}
