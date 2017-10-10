

using System;
using System.Collections.Generic;
using b2xtranslator.PptFileFormat;
using b2xtranslator.CommonTranslatorLib;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.OfficeDrawing;
using b2xtranslator.Tools;

namespace b2xtranslator.PresentationMLMapping
{
    class CharacterRunPropsMapping :
        AbstractOpenXmlMapping
    {
        protected ConversionContext _ctx;

        public CharacterRunPropsMapping(ConversionContext ctx, XmlWriter writer)
            : base(writer)
        {
            this._ctx = ctx;
        }

        public void Apply(CharacterRun run, string startElement, RegularContainer slide, ref string lastColor, ref string lastSize, ref string lastTypeface, string lang, string altLang, TextMasterStyleAtom defaultStyle, int lvl, List<MouseClickInteractiveInfoContainer> mciics, ShapeTreeMapping parentShapeTreeMapping, uint position, bool insideTable)
        {

            this._writer.WriteStartElement("a", startElement, OpenXmlNamespaces.DrawingML);


            if (lang.Length == 0)
            {
                var siea = this._ctx.Ppt.DocumentRecord.FirstDescendantWithType<TextSIExceptionAtom>();
                if (siea != null)
                {
                    if (siea.si.lang)
                    {
                        switch (siea.si.lid)
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
                                    lang = System.Globalization.CultureInfo.GetCultureInfo(siea.si.lid).IetfLanguageTag;
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

            if (altLang.Length == 0)
            {
                var siea = this._ctx.Ppt.DocumentRecord.FirstDescendantWithType<TextSIExceptionAtom>();
                if (siea != null)
                {
                    if (siea.si.altLang)
                    {
                        switch (siea.si.altLid)
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
                                    altLang = System.Globalization.CultureInfo.GetCultureInfo(siea.si.altLid).IetfLanguageTag;
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

            if (lang.Length > 0)
                this._writer.WriteAttributeString("lang", lang);

            if (altLang.Length > 0)
                this._writer.WriteAttributeString("altLang", altLang);

            bool runExists = run != null;

            if (runExists && run.SizePresent)
            {
                if (run.Size > 0)
                {
                    this._writer.WriteAttributeString("sz", (run.Size * 100).ToString());
                    lastSize = (run.Size * 100).ToString();
                }
            }
            else if (lastSize.Length > 0)
            {
                this._writer.WriteAttributeString("sz", lastSize);
            }
            else if (defaultStyle != null)
            {
                if (defaultStyle.CRuns[lvl].SizePresent)
                {
                    this._writer.WriteAttributeString("sz", (defaultStyle.CRuns[lvl].Size * 100).ToString());
                }
            }

            if (runExists && run.StyleFlagsFieldPresent)
            {
                if ((run.Style & StyleMask.IsBold) == StyleMask.IsBold) this._writer.WriteAttributeString("b", "1");
                if ((run.Style & StyleMask.IsItalic) == StyleMask.IsItalic) this._writer.WriteAttributeString("i", "1");
                if ((run.Style & StyleMask.IsUnderlined) == StyleMask.IsUnderlined) this._writer.WriteAttributeString("u", "sng");
            }
            else if (defaultStyle != null && defaultStyle.CRuns[lvl].StyleFlagsFieldPresent)
            {
                if ((defaultStyle.CRuns[lvl].Style & StyleMask.IsBold) == StyleMask.IsBold) this._writer.WriteAttributeString("b", "1");
                if ((defaultStyle.CRuns[lvl].Style & StyleMask.IsItalic) == StyleMask.IsItalic) this._writer.WriteAttributeString("i", "1");
                if ((defaultStyle.CRuns[lvl].Style & StyleMask.IsUnderlined) == StyleMask.IsUnderlined) this._writer.WriteAttributeString("u", "sng");
            }

            if (runExists && run.ColorPresent)
            {
                writeSolidFill(slide, run, ref lastColor);
            }
            else if (lastColor.Length > 0)
            {
                this._writer.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingML);
                this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("val", lastColor);
                this._writer.WriteEndElement();
                this._writer.WriteEndElement();
            }
            else if (defaultStyle != null)
            {
                if (defaultStyle.CRuns[lvl].ColorPresent)
                {
                    writeSolidFill((RegularContainer)defaultStyle.ParentRecord, defaultStyle.CRuns[lvl], ref lastColor);
                }
            }

            if (runExists && run.StyleFlagsFieldPresent)
            {
                if ((run.Style & StyleMask.HasShadow) == StyleMask.HasShadow)
                {
                    //TODO: these values are default and have to be replaced
                    this._writer.WriteStartElement("a", "effectLst", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteStartElement("a", "outerShdw", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("blurRad", "38100");
                    this._writer.WriteAttributeString("dist", "38100");
                    this._writer.WriteAttributeString("dir", "2700000");
                    this._writer.WriteAttributeString("algn", "tl");
                    this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("val", "C0C0C0");
                    this._writer.WriteEndElement();
                    this._writer.WriteEndElement();
                    this._writer.WriteEndElement();
                }

                if ((run.Style & StyleMask.IsEmbossed) == StyleMask.IsEmbossed)
                {
                    //TODO: these values are default and have to be replaced
                    this._writer.WriteStartElement("a", "effectDag", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("name", "");
                    this._writer.WriteStartElement("a", "cont", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("type", "tree");
                    this._writer.WriteAttributeString("name", "");
                    this._writer.WriteStartElement("a", "effect", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("ref", "fillLine");
                    this._writer.WriteEndElement();
                    this._writer.WriteStartElement("a", "outerShdw", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("dist", "38100");
                    this._writer.WriteAttributeString("dir", "13500000");
                    this._writer.WriteAttributeString("algn", "br");
                    this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("val", "FFFFFF");
                    this._writer.WriteEndElement();
                    this._writer.WriteEndElement();
                    this._writer.WriteEndElement();
                    this._writer.WriteStartElement("a", "cont", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("type", "tree");
                    this._writer.WriteAttributeString("name", "");
                    this._writer.WriteStartElement("a", "effect", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("ref", "fillLine");
                    this._writer.WriteEndElement();
                    this._writer.WriteStartElement("a", "outerShdw", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("dist", "38100");
                    this._writer.WriteAttributeString("dir", "2700000");
                    this._writer.WriteAttributeString("algn", "tl");
                    this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("val", "999999");
                    this._writer.WriteEndElement();
                    this._writer.WriteEndElement();
                    this._writer.WriteEndElement();
                    this._writer.WriteStartElement("a", "effect", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("ref", "fillLine");
                    this._writer.WriteEndElement();
                    this._writer.WriteEndElement();
                }

                //TODOS
                //HasAsianSmartQuotes 
                //HasHorizonNumRendering 
                //ExtensionNibble 

            }

            //TODOs:
            //run.ANSITypefacePresent
            //run.FEOldTypefacePresent
            //run.PositionPresent
            //run.SymbolTypefacePresent
            //run.TypefacePresent

            if (runExists && run.TypefacePresent)
            {
                this._writer.WriteStartElement("a", "latin", OpenXmlNamespaces.DrawingML);
                try
                {
                    var fonts = this._ctx.Ppt.DocumentRecord.FirstChildWithType<b2xtranslator.PptFileFormat.Environment>().FirstChildWithType<FontCollection>();
                    var entity = fonts.entities[(int)run.TypefaceIdx];
                    if (entity.TypeFace.IndexOf('\0') > 0)
                    {
                        this._writer.WriteAttributeString("typeface", entity.TypeFace.Substring(0, entity.TypeFace.IndexOf('\0')));
                        lastTypeface = entity.TypeFace.Substring(0, entity.TypeFace.IndexOf('\0'));
                    }
                    else
                    {
                        this._writer.WriteAttributeString("typeface", entity.TypeFace);
                        lastTypeface = entity.TypeFace;
                    }
                    //_writer.WriteAttributeString("charset", "0");
                }
                catch (Exception)
                {
                    throw;
                }

                this._writer.WriteEndElement();
            }
            else if (lastTypeface.Length > 0)
            {
                this._writer.WriteStartElement("a", "latin", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("typeface", lastTypeface);
                this._writer.WriteEndElement();
            }
            else if (defaultStyle != null && defaultStyle.CRuns[lvl].TypefacePresent)
            {
                this._writer.WriteStartElement("a", "latin", OpenXmlNamespaces.DrawingML);

                var fonts = this._ctx.Ppt.DocumentRecord.FirstChildWithType<b2xtranslator.PptFileFormat.Environment>().FirstChildWithType<FontCollection>();
                var entity = fonts.entities[(int)defaultStyle.CRuns[lvl].TypefaceIdx];
                if (entity.TypeFace.IndexOf('\0') > 0)
                {
                    this._writer.WriteAttributeString("typeface", entity.TypeFace.Substring(0, entity.TypeFace.IndexOf('\0')));
                    lastTypeface = entity.TypeFace.Substring(0, entity.TypeFace.IndexOf('\0'));
                }
                else
                {
                    this._writer.WriteAttributeString("typeface", entity.TypeFace);
                    lastTypeface = entity.TypeFace;
                }
                //_writer.WriteAttributeString("charset", "0");

                this._writer.WriteEndElement();
            }
            else
            {
                if (insideTable)
                    if (slide.FirstChildWithType<SlideAtom>() != null && this._ctx.Ppt.FindMasterRecordById(slide.FirstChildWithType<SlideAtom>().MasterId) != null)
                        foreach (var item in this._ctx.Ppt.FindMasterRecordById(slide.FirstChildWithType<SlideAtom>().MasterId).AllChildrenWithType<TextMasterStyleAtom>())
                        {
                            if (item.Instance == 1)
                            {
                                if (item.CRuns.Count > 0 && item.CRuns[0].TypefacePresent)
                                {
                                    this._writer.WriteStartElement("a", "latin", OpenXmlNamespaces.DrawingML);

                                    var fonts = this._ctx.Ppt.DocumentRecord.FirstChildWithType<b2xtranslator.PptFileFormat.Environment>().FirstChildWithType<FontCollection>();
                                    var entity = fonts.entities[(int)item.CRuns[0].TypefaceIdx];
                                    if (entity.TypeFace.IndexOf('\0') > 0)
                                    {
                                        this._writer.WriteAttributeString("typeface", entity.TypeFace.Substring(0, entity.TypeFace.IndexOf('\0')));
                                        lastTypeface = entity.TypeFace.Substring(0, entity.TypeFace.IndexOf('\0'));
                                    }
                                    else
                                    {
                                        this._writer.WriteAttributeString("typeface", entity.TypeFace);
                                        lastTypeface = entity.TypeFace;
                                    }
                                    //_writer.WriteAttributeString("charset", "0");

                                    this._writer.WriteEndElement();
                                }
                            }
                        }
                //        try
                //        {
                //            CharacterRun cr = _ctx.Ppt.DocumentRecord.FirstChildWithType<PptFileFormat.Environment>().FirstChildWithType<TextMasterStyleAtom>().CRuns[0];
                //            if (cr.TypefacePresent)
                //            {
                //                    FontCollection fonts = _ctx.Ppt.DocumentRecord.FirstChildWithType<b2xtranslator.PptFileFormat.Environment>().FirstChildWithType<FontCollection>();
                //                    FontEntityAtom entity = fonts.entities[(int)cr.TypefaceIdx];
                //                    if (entity.TypeFace.IndexOf('\0') > 0)
                //                    {
                //                        _writer.WriteStartElement("a", "latin", OpenXmlNamespaces.DrawingML);
                //                        _writer.WriteAttributeString("typeface", entity.TypeFace.Substring(0, entity.TypeFace.IndexOf('\0')));
                //                        _writer.WriteEndElement();
                //                    }
                //                    else
                //                    {
                //                        _writer.WriteStartElement("a", "latin", OpenXmlNamespaces.DrawingML);
                //                        _writer.WriteAttributeString("typeface", entity.TypeFace);
                //                        _writer.WriteEndElement();
                //                    }
                //                }
                //        }
                //        catch (Exception ex)
                //        {
                //            //throw;
                //        }

            }

            if (runExists && run.FEOldTypefacePresent)
            {
                try
                {
                    var fonts = this._ctx.Ppt.DocumentRecord.FirstChildWithType<b2xtranslator.PptFileFormat.Environment>().FirstChildWithType<FontCollection>();
                    var entity = fonts.entities[(int)run.FEOldTypefaceIdx];
                    if (entity.TypeFace.IndexOf('\0') > 0)
                    {
                        this._writer.WriteStartElement("a", "ea", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("typeface", entity.TypeFace.Substring(0, entity.TypeFace.IndexOf('\0')));
                        this._writer.WriteEndElement();
                    }
                    else
                    {
                        this._writer.WriteStartElement("a", "ea", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("typeface", entity.TypeFace);
                        this._writer.WriteEndElement();
                    }
                }
                catch (Exception)
                {
                    //throw;
                }
            }
            else
            {
                try
                {
                    var cr = this._ctx.Ppt.DocumentRecord.FirstChildWithType<PptFileFormat.Environment>().FirstChildWithType<TextMasterStyleAtom>().CRuns[0];
                    if (cr.FEOldTypefacePresent)
                    {
                        var fonts = this._ctx.Ppt.DocumentRecord.FirstChildWithType<b2xtranslator.PptFileFormat.Environment>().FirstChildWithType<FontCollection>();
                        var entity = fonts.entities[(int)cr.FEOldTypefaceIdx];
                        if (entity.TypeFace.IndexOf('\0') > 0)
                        {
                            this._writer.WriteStartElement("a", "ea", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("typeface", entity.TypeFace.Substring(0, entity.TypeFace.IndexOf('\0')));
                            this._writer.WriteEndElement();
                        }
                        else
                        {
                            this._writer.WriteStartElement("a", "ea", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("typeface", entity.TypeFace);
                            this._writer.WriteEndElement();
                        }
                    }
                }
                catch (Exception)
                {
                    //throw;
                }

            }

            if (runExists && run.SymbolTypefacePresent)
            {

                try
                {
                    var fonts = this._ctx.Ppt.DocumentRecord.FirstChildWithType<b2xtranslator.PptFileFormat.Environment>().FirstChildWithType<FontCollection>();
                    var entity = fonts.entities[(int)run.SymbolTypefaceIdx];
                    if (entity.TypeFace.IndexOf('\0') > 0)
                    {
                        this._writer.WriteStartElement("a", "sym", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("typeface", entity.TypeFace.Substring(0, entity.TypeFace.IndexOf('\0')));
                        this._writer.WriteEndElement();
                    }
                    else
                    {
                        this._writer.WriteStartElement("a", "sym", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("typeface", entity.TypeFace);
                        this._writer.WriteEndElement();
                    }
                }
                catch (Exception)
                {
                    //throw;
                }


            }

            if (mciics != null && mciics.Count > 0)
            {
                foreach (var mciic in mciics)
                {

                    var iia = mciic.FirstChildWithType<InteractiveInfoAtom>();
                    var tiia = mciic.Range;

                    if (tiia.begin <= position && tiia.end > position)
                        if (iia != null)
                        {
                            if (iia.action == InteractiveInfoActionEnum.Hyperlink)
                            {
                                foreach (var c in this._ctx.Ppt.DocumentRecord.FirstDescendantWithType<ExObjListContainer>().AllChildrenWithType<ExHyperlinkContainer>())
                                {
                                    var a = c.FirstChildWithType<ExHyperlinkAtom>();
                                    if (a.exHyperlinkId == iia.exHyperlinkIdRef)
                                    {
                                        var s = c.FirstChildWithType<CStringAtom>();
                                        var er = parentShapeTreeMapping.parentSlideMapping.targetPart.AddExternalRelationship(OpenXmlRelationshipTypes.HyperLink, s.Text);

                                        this._writer.WriteStartElement("a", "hlinkClick", OpenXmlNamespaces.DrawingML);
                                        this._writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, er.Id.ToString());
                                        this._writer.WriteEndElement();
                                    }
                                }

                            }
                        }
                }
            }

            this._writer.WriteEndElement();
        }


        public void writeSolidFill(RegularContainer slide, CharacterRun run, ref string lastColor)
        {
            this._writer.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingML);

            if (run.Color.IsSchemeColor) //TODO: to be fully implemented
            {
                //_writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);

                if (slide == null)
                {
                    ////TODO: what shall be used in this case (happens for default text style in presentation.xml)
                    //_writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    //_writer.WriteAttributeString("val", "000000");
                    //_writer.WriteEndElement();

                    this._writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);
                    switch (run.Color.Index)
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
                            lastColor = run.Color.Red.ToString("X").PadLeft(2, '0') + run.Color.Green.ToString("X").PadLeft(2, '0') + run.Color.Blue.ToString("X").PadLeft(2, '0');
                            this._writer.WriteAttributeString("val", lastColor);
                            break;
                        case 0xFF: //undefined
                            break;
                    }
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
                    foreach (var color in colors)
                    {
                        if (color.Instance == 1) MasterScheme = color;
                    }

                    this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    switch (run.Color.Index)
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
                            lastColor = run.Color.Red.ToString("X").PadLeft(2, '0') + run.Color.Green.ToString("X").PadLeft(2, '0') + run.Color.Blue.ToString("X").PadLeft(2, '0');
                            this._writer.WriteAttributeString("val", lastColor);
                            break;
                        case 0xFF: //undefined
                            break;
                    }
                    this._writer.WriteEndElement();
                    //_writer.WriteEndElement();
                }
            }
            else
            {
                this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                lastColor = run.Color.Red.ToString("X").PadLeft(2, '0') + run.Color.Green.ToString("X").PadLeft(2, '0') + run.Color.Blue.ToString("X").PadLeft(2, '0');
                this._writer.WriteAttributeString("val", lastColor);
                this._writer.WriteEndElement();
            }
            this._writer.WriteEndElement();
        }

    }
}
