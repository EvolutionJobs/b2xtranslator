/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.DocFileFormat;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
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
            _parentDoc = parentDoc;
            _pPr = _nodeFactory.CreateElement("w", "pPr", OpenXmlNamespaces.WordprocessingML);
            _framePr = _nodeFactory.CreateElement("w", "framePr", OpenXmlNamespaces.WordprocessingML);
            _paraEndChpx = paraEndChpx;
            _ctx = ctx;
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
            _parentDoc = parentDoc;
            _pPr = _nodeFactory.CreateElement("w", "pPr", OpenXmlNamespaces.WordprocessingML);
            _framePr = _nodeFactory.CreateElement("w", "framePr", OpenXmlNamespaces.WordprocessingML);
            _paraEndChpx = paraEndChpx;
            _sepx = sepx;
            _ctx = ctx;
            _sectionNr = sectionNr;
        }

        public void Apply(ParagraphPropertyExceptions papx)
        {
            XmlElement ind = _nodeFactory.CreateElement("w", "ind", OpenXmlNamespaces.WordprocessingML);
            XmlElement numPr = _nodeFactory.CreateElement("w", "numPr", OpenXmlNamespaces.WordprocessingML);
            XmlElement pBdr = _nodeFactory.CreateElement("w", "pBdr", OpenXmlNamespaces.WordprocessingML);
            XmlElement spacing = _nodeFactory.CreateElement("w", "spacing", OpenXmlNamespaces.WordprocessingML);
            XmlElement jc = null;

            //append style id , do not append "Normal" style (istd 0)
            XmlElement pStyle = _nodeFactory.CreateElement("w", "pStyle", OpenXmlNamespaces.WordprocessingML);
            XmlAttribute styleId = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
            styleId.Value = StyleSheetMapping.MakeStyleId(_parentDoc.Styles.Styles[papx.istd]);
            pStyle.Attributes.Append(styleId);
            _pPr.AppendChild(pStyle);

            //append formatting of paragraph end mark
            if (_paraEndChpx != null)
            {
                XmlElement rPr = _nodeFactory.CreateElement("w", "rPr", OpenXmlNamespaces.WordprocessingML);
                
                //append properties
                _paraEndChpx.Convert(new CharacterPropertiesMapping(rPr, _ctx.Doc, new RevisionData(_paraEndChpx), papx, false));

                var rev = new RevisionData(_paraEndChpx);
                //append delete infos
                if (rev.Type == RevisionData.RevisionType.Deleted)
                {
                    XmlElement del = _nodeFactory.CreateElement("w", "del", OpenXmlNamespaces.WordprocessingML);
                    rPr.AppendChild(del);
                }

                if(rPr.ChildNodes.Count >0 )
                {
                    _pPr.AppendChild(rPr);
                }
            }

            bool isRightToLeft = false;
            foreach (SinglePropertyModifier sprm in papx.grpprl)
            {
                switch (sprm.OpCode)
                {
                    //rsid for paragraph property enditing (write to parent element)
                    case SinglePropertyModifier.OperationCode.sprmPRsid:
                        string rsid = String.Format("{0:x8}", System.BitConverter.ToInt32(sprm.Arguments, 0));
                        _ctx.AddRsid(rsid);
                        _writer.WriteAttributeString("w", "rsidP", OpenXmlNamespaces.WordprocessingML, rsid);
                        break;

                    //attributes
                    case SinglePropertyModifier.OperationCode.sprmPIpgp:
                        XmlAttribute divId = _nodeFactory.CreateAttribute("w", "divId", OpenXmlNamespaces.WordprocessingML);
                        divId.Value = System.BitConverter.ToUInt32(sprm.Arguments, 0).ToString();
                        _pPr.Attributes.Append(divId);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFAutoSpaceDE:
                        appendFlagAttribute(_pPr, sprm, "autoSpaceDE");
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFAutoSpaceDN:
                        appendFlagAttribute(_pPr, sprm, "autoSpaceDN");
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFContextualSpacing:
                        appendFlagAttribute(_pPr, sprm, "contextualSpacing");
                        break;
                    
                    //element flags
                    case SinglePropertyModifier.OperationCode.sprmPFBiDi:
                        isRightToLeft = true;
                        appendFlagElement(_pPr, sprm, "bidi", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFKeep:
                        appendFlagElement(_pPr, sprm, "keepLines", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFKeepFollow:
                        appendFlagElement(_pPr, sprm, "keepNext", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFKinsoku:
                        appendFlagElement(_pPr, sprm, "kinsoku", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFOverflowPunct:
                        appendFlagElement(_pPr, sprm, "overflowPunct", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFPageBreakBefore:
                        appendFlagElement(_pPr, sprm, "pageBreakBefore", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFNoAutoHyph:
                        appendFlagElement(_pPr, sprm, "su_pPressAutoHyphens", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFNoLineNumb:
                        appendFlagElement(_pPr, sprm, "su_pPressLineNumbers", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFNoAllowOverlap:
                        appendFlagElement(_pPr, sprm, "su_pPressOverlap", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFTopLinePunct:
                        appendFlagElement(_pPr, sprm, "topLinePunct", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFWidowControl:
                        appendFlagElement(_pPr, sprm, "widowControl", true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFWordWrap:
                        appendFlagElement(_pPr, sprm, "wordWrap", true);
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
                        var flValue = System.BitConverter.ToInt16(sprm.Arguments, 0);
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
                        XmlAttribute before = _nodeFactory.CreateAttribute("w", "before", OpenXmlNamespaces.WordprocessingML);
                        before.Value = System.BitConverter.ToUInt16(sprm.Arguments, 0).ToString();
                        spacing.Attributes.Append(before);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPDyaAfter:
                        XmlAttribute after = _nodeFactory.CreateAttribute("w", "after", OpenXmlNamespaces.WordprocessingML);
                        after.Value = System.BitConverter.ToUInt16(sprm.Arguments, 0).ToString();
                        spacing.Attributes.Append(after);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFDyaAfterAuto:
                        XmlAttribute afterAutospacing = _nodeFactory.CreateAttribute("w", "afterAutospacing", OpenXmlNamespaces.WordprocessingML);
                        afterAutospacing.Value = sprm.Arguments[0].ToString();
                        spacing.Attributes.Append(afterAutospacing);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPFDyaBeforeAuto:
                        XmlAttribute beforeAutospacing = _nodeFactory.CreateAttribute("w", "beforeAutospacing", OpenXmlNamespaces.WordprocessingML);
                        beforeAutospacing.Value = sprm.Arguments[0].ToString();
                        spacing.Attributes.Append(beforeAutospacing);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPDyaLine:
                        var lspd = new LineSpacingDescriptor(sprm.Arguments);
                        XmlAttribute line = _nodeFactory.CreateAttribute("w", "line", OpenXmlNamespaces.WordprocessingML);
                        line.Value = Math.Abs(lspd.dyaLine).ToString();
                        spacing.Attributes.Append(line);
                        XmlAttribute lineRule = _nodeFactory.CreateAttribute("w", "lineRule", OpenXmlNamespaces.WordprocessingML);
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
                        jc = _nodeFactory.CreateElement("w", "jc", OpenXmlNamespaces.WordprocessingML);
                        XmlAttribute jcVal = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                        jcVal.Value = ((Global.JustificationCode)sprm.Arguments[0]).ToString();
                        jc.Attributes.Append(jcVal);
                        break;

                    //borders
                    //case 0x461C:
                    case SinglePropertyModifier.OperationCode.sprmPBrcTop:
                    //case 0x4424:
                    case SinglePropertyModifier.OperationCode.sprmPBrcTop80:
                        XmlNode topBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "top", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), topBorder);
                        addOrSetBorder(pBdr, topBorder);
                        break;
                    //case 0x461D:
                    case SinglePropertyModifier.OperationCode.sprmPBrcLeft:
                    //case 0x4425:
                    case SinglePropertyModifier.OperationCode.sprmPBrcLeft80:
                        XmlNode leftBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "left", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), leftBorder);
                        addOrSetBorder(pBdr, leftBorder);
                        break;
                    //case 0x461E:
                    case SinglePropertyModifier.OperationCode.sprmPBrcBottom:
                    //case 0x4426:
                    case SinglePropertyModifier.OperationCode.sprmPBrcBottom80:
                        XmlNode bottomBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "bottom", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), bottomBorder);
                        addOrSetBorder(pBdr, bottomBorder);
                        break;
                    //case 0x461F:
                    case SinglePropertyModifier.OperationCode.sprmPBrcRight:
                    //case 0x4427:
                    case SinglePropertyModifier.OperationCode.sprmPBrcRight80:
                        XmlNode rightBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "right", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), rightBorder);
                        addOrSetBorder(pBdr, rightBorder);
                        break;
                    //case 0x4620:
                    case SinglePropertyModifier.OperationCode.sprmPBrcBetween:
                    //case 0x4428:
                    case SinglePropertyModifier.OperationCode.sprmPBrcBetween80:
                        XmlNode betweenBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "between", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), betweenBorder);
                        addOrSetBorder(pBdr, betweenBorder);
                        break;
                    //case 0x4621:
                    case SinglePropertyModifier.OperationCode.sprmPBrcBar:
                    //case 0x4629:
                    case SinglePropertyModifier.OperationCode.sprmPBrcBar80:
                        XmlNode barBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "bar", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), barBorder);
                        addOrSetBorder(pBdr, barBorder);
                        break;
                    
                    //shading
                    case SinglePropertyModifier.OperationCode.sprmPShd80:
                    case SinglePropertyModifier.OperationCode.sprmPShd:
                        var desc = new ShadingDescriptor(sprm.Arguments);
                        appendShading(_pPr, desc);
                        break;
                    
                    //numbering
                    case SinglePropertyModifier.OperationCode.sprmPIlvl:
                        appendValueElement(numPr, "ilvl", sprm.Arguments[0].ToString(), true);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPIlfo:
                        var val  = System.BitConverter.ToUInt16(sprm.Arguments, 0);
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
                        XmlElement tabs = _nodeFactory.CreateElement("w", "tabs", OpenXmlNamespaces.WordprocessingML);
                        int pos = 0;
                        //read the removed tabs
                        byte itbdDelMax = sprm.Arguments[pos];
                        pos++;
                        for(int i=0; i<itbdDelMax; i++)
                        {
                            XmlElement tab = _nodeFactory.CreateElement("w", "tab", OpenXmlNamespaces.WordprocessingML);
                            //clear
                            XmlAttribute tabsVal = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                            tabsVal.Value = "clear";
                            tab.Attributes.Append(tabsVal);
                            //position
                            XmlAttribute tabsPos = _nodeFactory.CreateAttribute("w", "pos", OpenXmlNamespaces.WordprocessingML);
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
                            XmlElement tab = _nodeFactory.CreateElement("w", "tab", OpenXmlNamespaces.WordprocessingML);
                            //justification
                            XmlAttribute tabsVal = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                            tabsVal.Value = ((Global.JustificationCode)tbd.jc).ToString();
                            tab.Attributes.Append(tabsVal);
                            //tab leader type
                            XmlAttribute leader = _nodeFactory.CreateAttribute("w", "leader", OpenXmlNamespaces.WordprocessingML);
                            leader.Value = ((Global.TabLeader)tbd.tlc).ToString();
                            tab.Attributes.Append(leader);
                            //position
                            XmlAttribute tabsPos = _nodeFactory.CreateAttribute("w", "pos", OpenXmlNamespaces.WordprocessingML);
                            tabsPos.Value = System.BitConverter.ToInt16(sprm.Arguments, pos + (i * 2)).ToString();
                            tab.Attributes.Append(tabsPos);
                            tabs.AppendChild(tab);
                        }
                        _pPr.AppendChild(tabs);
                        break;

                    //frame properties

                    case SinglePropertyModifier.OperationCode.sprmPPc:
                        //position code
                        byte flag = sprm.Arguments[0];
                        var pcVert = (Global.VerticalPositionCode)((flag & 0x30) >> 4);
                        var pcHorz = (Global.HorizontalPositionCode)((flag & 0xC0) >> 6);
                        appendValueAttribute(_framePr, "hAnchor", pcHorz.ToString());
                        appendValueAttribute(_framePr, "vAnchor", pcVert.ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPWr:
                        var wrapping = (Global.TextFrameWrapping)sprm.Arguments[0];
                        appendValueAttribute(_framePr, "wrap", wrapping.ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPDxaAbs:
                        var frameX = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        appendValueAttribute(_framePr, "x", frameX.ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPDyaAbs:
                        var frameY = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        appendValueAttribute(_framePr, "y", frameY.ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPWHeightAbs:
                        var frameHeight = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        appendValueAttribute(_framePr, "h", frameHeight.ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPDxaWidth:
                        var frameWidth = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        appendValueAttribute(_framePr, "w", frameWidth.ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPDxaFromText:
                        var frameSpaceH = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        appendValueAttribute(_framePr, "hSpace", frameSpaceH.ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmPDyaFromText:
                        var frameSpaceV = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        appendValueAttribute(_framePr, "vSpace", frameSpaceV.ToString());
                        break;

                    //outline level
                    case SinglePropertyModifier.OperationCode.sprmPOutLvl:
                        appendValueElement(_pPr, "outlineLvl", sprm.Arguments[0].ToString(), false);
                        break;

                    default:
                        break;
                }
            }

            //append frame properties
            if (_framePr.Attributes.Count > 0)
            {
                _pPr.AppendChild(_framePr);
            }

            //append section properties
            if (_sepx != null)
            {
                XmlElement sectPr = _nodeFactory.CreateElement("w", "sectPr", OpenXmlNamespaces.WordprocessingML);
                _sepx.Convert(new SectionPropertiesMapping(sectPr, _ctx, _sectionNr));
                _pPr.AppendChild(sectPr);
            }

            //append indent
            if (ind.Attributes.Count > 0)
                _pPr.AppendChild(ind);

            //append spacing
            if (spacing.Attributes.Count > 0)
                _pPr.AppendChild(spacing);

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
                    _pPr.AppendChild(jc);
                }
            }

            //append numPr
            if (numPr.ChildNodes.Count > 0)
                _pPr.AppendChild(numPr);

            //append borders
            if (pBdr.ChildNodes.Count > 0)
                _pPr.AppendChild(pBdr);

            //write Properties
            if (_pPr.ChildNodes.Count > 0 || _pPr.Attributes.Count > 0)
            {
                _pPr.WriteTo(_writer);
            }
        }

        private bool containsLvlReference(List<SinglePropertyModifier> sprms)
        {
            bool ret = false;
            foreach (SinglePropertyModifier sprm in sprms)
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
            StyleSheetDescription style = _parentDoc.Styles.Styles[istd];
            foreach (SinglePropertyModifier sprm in style.papx.grpprl)
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
