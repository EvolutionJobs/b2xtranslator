using System;
using System.Collections.Generic;
using b2xtranslator.CommonTranslatorLib;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.DocFileFormat;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class ParagraphPropertiesMapping : PropertiesMapping,
          IMapping<ParagraphPropertyExceptions>
    {
        private ConversionContext _ctx;
        private XmlElement _pPr;
        private XmlElement _framePr;
        private SectionPropertyExceptions _sepx;
        private CharacterPropertyExceptions _paraEndChpx;
        private int _sectionNr;
        private WordDocument _parentDoc;

        public ParagraphPropertiesMapping(
            XmlWriter writer, 
            ConversionContext ctx, 
            WordDocument parentDoc,
            CharacterPropertyExceptions paraEndChpx)
            : base(writer)
        {
            this._parentDoc = parentDoc;
            this._pPr = this._nodeFactory.CreateElement("w", "pPr", OpenXmlNamespaces.WordprocessingML);
            this._framePr = this._nodeFactory.CreateElement("w", "framePr", OpenXmlNamespaces.WordprocessingML);
            this._paraEndChpx = paraEndChpx;
            this._ctx = ctx;
        }

        public ParagraphPropertiesMapping(
            XmlWriter writer, 
            ConversionContext ctx,
            WordDocument parentDoc,
            CharacterPropertyExceptions paraEndChpx, 
            SectionPropertyExceptions sepx,
            int sectionNr)
            : base(writer)
        {
            this._parentDoc = parentDoc;
            this._pPr = this._nodeFactory.CreateElement("w", "pPr", OpenXmlNamespaces.WordprocessingML);
            this._framePr = this._nodeFactory.CreateElement("w", "framePr", OpenXmlNamespaces.WordprocessingML);
            this._paraEndChpx = paraEndChpx;
            this._sepx = sepx;
            this._ctx = ctx;
            this._sectionNr = sectionNr;
        }

        public void Apply(ParagraphPropertyExceptions papx)
        {
            var ind = this._nodeFactory.CreateElement("w", "ind", OpenXmlNamespaces.WordprocessingML);
            var numPr = this._nodeFactory.CreateElement("w", "numPr", OpenXmlNamespaces.WordprocessingML);
            var pBdr = this._nodeFactory.CreateElement("w", "pBdr", OpenXmlNamespaces.WordprocessingML);
            var spacing = this._nodeFactory.CreateElement("w", "spacing", OpenXmlNamespaces.WordprocessingML);
            XmlElement jc = null;

            //append style id , do not append "Normal" style (istd 0)
            var pStyle = this._nodeFactory.CreateElement("w", "pStyle", OpenXmlNamespaces.WordprocessingML);
            var styleId = this._nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
            styleId.Value = StyleSheetMapping.MakeStyleId(this._parentDoc.Styles.Styles[papx.istd]);
            pStyle.Attributes.Append(styleId);
            this._pPr.AppendChild(pStyle);

            //append formatting of paragraph end mark
            if (this._paraEndChpx != null)
            {
                var rPr = this._nodeFactory.CreateElement("w", "rPr", OpenXmlNamespaces.WordprocessingML);

                //append properties
                this._paraEndChpx.Convert(new CharacterPropertiesMapping(rPr, this._ctx.Doc, new RevisionData(this._paraEndChpx), papx, false));

                var rev = new RevisionData(this._paraEndChpx);
                //append delete infos
                if (rev.Type == RevisionData.RevisionType.Deleted)
                {
                    var del = this._nodeFactory.CreateElement("w", "del", OpenXmlNamespaces.WordprocessingML);
                    rPr.AppendChild(del);
                }

                if(rPr.ChildNodes.Count >0 )
                {
                    this._pPr.AppendChild(rPr);
                }
            }

            bool isRightToLeft = false;
            foreach (var sprm in papx.grpprl)
            {
                switch (sprm.OpCode)
                {
                    //rsid for paragraph property enditing (write to parent element)
                    case SinglePropertyModifier.OperationCode.sprmPRsid:
                        string rsid = string.Format("{0:x8}", System.BitConverter.ToInt32(sprm.Arguments, 0));
                        this._ctx.AddRsid(rsid);
                        this._writer.WriteAttributeString("w", "rsidP", OpenXmlNamespaces.WordprocessingML, rsid);
                        break;

                    //attributes
                    case SinglePropertyModifier.OperationCode.sprmPIpgp:
                        var divId = this._nodeFactory.CreateAttribute("w", "divId", OpenXmlNamespaces.WordprocessingML);
                        divId.Value = System.BitConverter.ToUInt32(sprm.Arguments, 0).ToString();
                        this._pPr.Attributes.Append(divId);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFAutoSpaceDE:
                        appendFlagAttribute(this._pPr, sprm, "autoSpaceDE");
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFAutoSpaceDN:
                        appendFlagAttribute(this._pPr, sprm, "autoSpaceDN");
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFContextualSpacing:
                        appendFlagAttribute(this._pPr, sprm, "contextualSpacing");
                        break;
                    
                    //element flags
                    case SinglePropertyModifier.OperationCode.sprmPFBiDi:
                        isRightToLeft = true;
                        appendFlagElement(this._pPr, sprm, "bidi", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFKeep:
                        appendFlagElement(this._pPr, sprm, "keepLines", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFKeepFollow:
                        appendFlagElement(this._pPr, sprm, "keepNext", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFKinsoku:
                        appendFlagElement(this._pPr, sprm, "kinsoku", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFOverflowPunct:
                        appendFlagElement(this._pPr, sprm, "overflowPunct", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFPageBreakBefore:
                        appendFlagElement(this._pPr, sprm, "pageBreakBefore", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFNoAutoHyph:
                        appendFlagElement(this._pPr, sprm, "su_pPressAutoHyphens", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFNoLineNumb:
                        appendFlagElement(this._pPr, sprm, "su_pPressLineNumbers", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFNoAllowOverlap:
                        appendFlagElement(this._pPr, sprm, "su_pPressOverlap", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFTopLinePunct:
                        appendFlagElement(this._pPr, sprm, "topLinePunct", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFWidowControl:
                        appendFlagElement(this._pPr, sprm, "widowControl", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFWordWrap:
                        appendFlagElement(this._pPr, sprm, "wordWrap", true);
                        break;

                    //indentation
                    case SinglePropertyModifier.OperationCode.sprmPDxaLeft:
                    case SinglePropertyModifier.OperationCode.sprmPDxaLeft80:
                    case SinglePropertyModifier.OperationCode.sprmPNest:
                    case SinglePropertyModifier.OperationCode.sprmPNest80:
                        appendValueAttribute(ind, "left", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPDxcLeft:
                        appendValueAttribute(ind, "leftChars", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPDxaLeft1:
                    case SinglePropertyModifier.OperationCode.sprmPDxaLeft180:
                        short flValue = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        string flName;
                        if (flValue >= 0)
                        {
                            flName = "firstLine";
                        }
                        else
                        {
                            flName = "hanging";
                            flValue *= -1;
                        }
                        appendValueAttribute(ind, flName, flValue.ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPDxcLeft1:
                        appendValueAttribute(ind, "firstLineChars", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPDxaRight:
                    case SinglePropertyModifier.OperationCode.sprmPDxaRight80:
                        appendValueAttribute(ind, "right", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPDxcRight:
                        appendValueAttribute(ind, "rightChars", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;

                    //spacing
                    case SinglePropertyModifier.OperationCode.sprmPDyaBefore:
                        var before = this._nodeFactory.CreateAttribute("w", "before", OpenXmlNamespaces.WordprocessingML);
                        before.Value = System.BitConverter.ToUInt16(sprm.Arguments, 0).ToString();
                        spacing.Attributes.Append(before);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPDyaAfter:
                        var after = this._nodeFactory.CreateAttribute("w", "after", OpenXmlNamespaces.WordprocessingML);
                        after.Value = System.BitConverter.ToUInt16(sprm.Arguments, 0).ToString();
                        spacing.Attributes.Append(after);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFDyaAfterAuto:
                        var afterAutospacing = this._nodeFactory.CreateAttribute("w", "afterAutospacing", OpenXmlNamespaces.WordprocessingML);
                        afterAutospacing.Value = sprm.Arguments[0].ToString();
                        spacing.Attributes.Append(afterAutospacing);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFDyaBeforeAuto:
                        var beforeAutospacing = this._nodeFactory.CreateAttribute("w", "beforeAutospacing", OpenXmlNamespaces.WordprocessingML);
                        beforeAutospacing.Value = sprm.Arguments[0].ToString();
                        spacing.Attributes.Append(beforeAutospacing);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPDyaLine:
                        var lspd = new LineSpacingDescriptor(sprm.Arguments);
                        var line = this._nodeFactory.CreateAttribute("w", "line", OpenXmlNamespaces.WordprocessingML);
                        line.Value = Math.Abs(lspd.dyaLine).ToString();
                        spacing.Attributes.Append(line);
                        var lineRule = this._nodeFactory.CreateAttribute("w", "lineRule", OpenXmlNamespaces.WordprocessingML);
                        if(!lspd.fMultLinespace && lspd.dyaLine < 0)
                            lineRule.Value = "exact";
                        else if(!lspd.fMultLinespace && lspd.dyaLine > 0)
                            lineRule.Value = "atLeast";
                        //no line rule means auto
                        spacing.Attributes.Append(lineRule);
                        break;

                    //justification code
                    case SinglePropertyModifier.OperationCode.sprmPJc:
                    case SinglePropertyModifier.OperationCode.sprmPJc80:
                        jc = this._nodeFactory.CreateElement("w", "jc", OpenXmlNamespaces.WordprocessingML);
                        var jcVal = this._nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                        jcVal.Value = ((Global.JustificationCode)sprm.Arguments[0]).ToString();
                        jc.Attributes.Append(jcVal);
                        break;

                    //borders
                    //case 0x461C:
                    case SinglePropertyModifier.OperationCode.sprmPBrcTop:
                    //case 0x4424:
                    case SinglePropertyModifier.OperationCode.sprmPBrcTop80:
                        var topBorder = this._nodeFactory.CreateNode(XmlNodeType.Element, "w", "top", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), topBorder);
                        addOrSetBorder(pBdr, topBorder);
                        break;
                    //case 0x461D:
                    case SinglePropertyModifier.OperationCode.sprmPBrcLeft:
                    //case 0x4425:
                    case SinglePropertyModifier.OperationCode.sprmPBrcLeft80:
                        var leftBorder = this._nodeFactory.CreateNode(XmlNodeType.Element, "w", "left", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), leftBorder);
                        addOrSetBorder(pBdr, leftBorder);
                        break;
                    //case 0x461E:
                    case SinglePropertyModifier.OperationCode.sprmPBrcBottom:
                    //case 0x4426:
                    case SinglePropertyModifier.OperationCode.sprmPBrcBottom80:
                        var bottomBorder = this._nodeFactory.CreateNode(XmlNodeType.Element, "w", "bottom", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), bottomBorder);
                        addOrSetBorder(pBdr, bottomBorder);
                        break;
                    //case 0x461F:
                    case SinglePropertyModifier.OperationCode.sprmPBrcRight:
                    //case 0x4427:
                    case SinglePropertyModifier.OperationCode.sprmPBrcRight80:
                        var rightBorder = this._nodeFactory.CreateNode(XmlNodeType.Element, "w", "right", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), rightBorder);
                        addOrSetBorder(pBdr, rightBorder);
                        break;
                    //case 0x4620:
                    case SinglePropertyModifier.OperationCode.sprmPBrcBetween:
                    //case 0x4428:
                    case SinglePropertyModifier.OperationCode.sprmPBrcBetween80:
                        var betweenBorder = this._nodeFactory.CreateNode(XmlNodeType.Element, "w", "between", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), betweenBorder);
                        addOrSetBorder(pBdr, betweenBorder);
                        break;
                    //case 0x4621:
                    case SinglePropertyModifier.OperationCode.sprmPBrcBar:
                    //case 0x4629:
                    case SinglePropertyModifier.OperationCode.sprmPBrcBar80:
                        var barBorder = this._nodeFactory.CreateNode(XmlNodeType.Element, "w", "bar", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), barBorder);
                        addOrSetBorder(pBdr, barBorder);
                        break;
                    
                    //shading
                    case SinglePropertyModifier.OperationCode.sprmPShd80:
                    case SinglePropertyModifier.OperationCode.sprmPShd:
                        var desc = new ShadingDescriptor(sprm.Arguments);
                        appendShading(this._pPr, desc);
                        break;
                    
                    //numbering
                    case SinglePropertyModifier.OperationCode.sprmPIlvl:
                        appendValueElement(numPr, "ilvl", sprm.Arguments[0].ToString(), true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPIlfo:
                        ushort val  = System.BitConverter.ToUInt16(sprm.Arguments, 0);
                        appendValueElement(numPr, "numId", val.ToString(), true);

                        ////check if there is a ilvl reference, if not, check the count of LVLs.
                        ////if only one LVL exists in the referenced list, create a hard reference to that LVL
                        //if (containsLvlReference(papx.grpprl) == false)
                        //{
                        //    ListFormatOverride lfo = _ctx.Doc.ListFormatOverrideTable[val];
                        //    int index = NumberingMapping.FindIndexbyId(_ctx.Doc.ListTable, lfo.lsid);
                        //    ListData lst = _ctx.Doc.ListTable[index];
                        //    if (lst.rglvl.Length == 1)
                        //    {
                        //        appendValueElement(numPr, "ilvl", "0", true);
                        //    }
                        //}
                        break;

                    //tabs
                    case SinglePropertyModifier.OperationCode.sprmPChgTabsPapx:
                    case SinglePropertyModifier.OperationCode.sprmPChgTabs:
                        var tabs = this._nodeFactory.CreateElement("w", "tabs", OpenXmlNamespaces.WordprocessingML);
                        int pos = 0;
                        //read the removed tabs
                        byte itbdDelMax = sprm.Arguments[pos];
                        pos++;
                        for(int i=0; i<itbdDelMax; i++)
                        {
                            var tab = this._nodeFactory.CreateElement("w", "tab", OpenXmlNamespaces.WordprocessingML);
                            //clear
                            var tabsVal = this._nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                            tabsVal.Value = "clear";
                            tab.Attributes.Append(tabsVal);
                            //position
                            var tabsPos = this._nodeFactory.CreateAttribute("w", "pos", OpenXmlNamespaces.WordprocessingML);
                            tabsPos.Value = System.BitConverter.ToInt16(sprm.Arguments, pos).ToString();
                            tab.Attributes.Append(tabsPos);
                            tabs.AppendChild(tab);
                            
                            //skip the tolerence array in sprm 0xC615
                            if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPChgTabs)
                                pos += 4;
                            else
                                pos += 2;
                        }
                        //read the added tabs
                        byte itbdAddMax = sprm.Arguments[pos];
                        pos++;
                        for (int i = 0; i < itbdAddMax; i++)
                        {
                            var tbd = new TabDescriptor(sprm.Arguments[pos + (itbdAddMax * 2) + i]);
                            var tab = this._nodeFactory.CreateElement("w", "tab", OpenXmlNamespaces.WordprocessingML);
                            //justification
                            var tabsVal = this._nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                            tabsVal.Value = ((Global.JustificationCode)tbd.jc).ToString();
                            tab.Attributes.Append(tabsVal);
                            //tab leader type
                            var leader = this._nodeFactory.CreateAttribute("w", "leader", OpenXmlNamespaces.WordprocessingML);
                            leader.Value = ((Global.TabLeader)tbd.tlc).ToString();
                            tab.Attributes.Append(leader);
                            //position
                            var tabsPos = this._nodeFactory.CreateAttribute("w", "pos", OpenXmlNamespaces.WordprocessingML);
                            tabsPos.Value = System.BitConverter.ToInt16(sprm.Arguments, pos + (i * 2)).ToString();
                            tab.Attributes.Append(tabsPos);
                            tabs.AppendChild(tab);
                        }
                        this._pPr.AppendChild(tabs);
                        break;

                    //frame properties

                    case SinglePropertyModifier.OperationCode.sprmPPc:
                        //position code
                        byte flag = sprm.Arguments[0];
                        var pcVert = (Global.VerticalPositionCode)((flag & 0x30) >> 4);
                        var pcHorz = (Global.HorizontalPositionCode)((flag & 0xC0) >> 6);
                        appendValueAttribute(this._framePr, "hAnchor", pcHorz.ToString());
                        appendValueAttribute(this._framePr, "vAnchor", pcVert.ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPWr:
                        var wrapping = (Global.TextFrameWrapping)sprm.Arguments[0];
                        appendValueAttribute(this._framePr, "wrap", wrapping.ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPDxaAbs:
                        short frameX = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        appendValueAttribute(this._framePr, "x", frameX.ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPDyaAbs:
                        short frameY = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        appendValueAttribute(this._framePr, "y", frameY.ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPWHeightAbs:
                        short frameHeight = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        appendValueAttribute(this._framePr, "h", frameHeight.ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPDxaWidth:
                        short frameWidth = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        appendValueAttribute(this._framePr, "w", frameWidth.ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPDxaFromText:
                        short frameSpaceH = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        appendValueAttribute(this._framePr, "hSpace", frameSpaceH.ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPDyaFromText:
                        short frameSpaceV = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        appendValueAttribute(this._framePr, "vSpace", frameSpaceV.ToString());
                        break;

                    //outline level
                    case SinglePropertyModifier.OperationCode.sprmPOutLvl:
                        appendValueElement(this._pPr, "outlineLvl", sprm.Arguments[0].ToString(), false);
                        break;

                    default:
                        break;
                }
            }

            //append frame properties
            if (this._framePr.Attributes.Count > 0)
            {
                this._pPr.AppendChild(this._framePr);
            }

            //append section properties
            if (this._sepx != null)
            {
                var sectPr = this._nodeFactory.CreateElement("w", "sectPr", OpenXmlNamespaces.WordprocessingML);
                this._sepx.Convert(new SectionPropertiesMapping(sectPr, this._ctx, this._sectionNr));
                this._pPr.AppendChild(sectPr);
            }

            //append indent
            if (ind.Attributes.Count > 0)
                this._pPr.AppendChild(ind);

            //append spacing
            if (spacing.Attributes.Count > 0)
                this._pPr.AppendChild(spacing);

            //append justification
            if (jc != null)
            {
                var jcVal = jc.Attributes["val", OpenXmlNamespaces.WordprocessingML];
                if ((isRightToLeft || isStyleRightToLeft(papx.istd)) && jcVal.Value == "right")
                {
                    //ignore jc="right" for RTL documents
                }
                else
                {
                    this._pPr.AppendChild(jc);
                }
            }

            //append numPr
            if (numPr.ChildNodes.Count > 0)
                this._pPr.AppendChild(numPr);

            //append borders
            if (pBdr.ChildNodes.Count > 0)
                this._pPr.AppendChild(pBdr);

            //write Properties
            if (this._pPr.ChildNodes.Count > 0 || this._pPr.Attributes.Count > 0)
            {
                this._pPr.WriteTo(this._writer);
            }
        }

        private bool containsLvlReference(List<SinglePropertyModifier> sprms)
        {
            bool ret = false;
            foreach (var sprm in sprms)
            {
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPIlvl)
                {
                    ret = true;
                    break;
                }
            }
            return ret;
        }

        private bool isStyleRightToLeft(ushort istd)
        {
            var style = this._parentDoc.Styles.Styles[istd];
            foreach (var sprm in style.papx.grpprl)
            {
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPFBiDi)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
