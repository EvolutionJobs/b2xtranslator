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
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class CharacterPropertiesMapping : PropertiesMapping,
          IMapping<CharacterPropertyExceptions>
    {
        private WordDocument _doc;
        private XmlElement _rPr;
        private ushort _currentIstd;
        private RevisionData _revisionData;
        private bool _styleChpx;
        private ParagraphPropertyExceptions _currentPapx;
        List<CharacterPropertyExceptions> _hierarchy;

        private enum SuperscriptIndex
        {
            baseline,
            superscript,
            subscript
        }

        public CharacterPropertiesMapping(XmlWriter writer, WordDocument doc, RevisionData rev, ParagraphPropertyExceptions currentPapx, bool styleChpx)
            : base(writer)
        {
            _doc = doc;
            _rPr = _nodeFactory.CreateElement("w", "rPr", OpenXmlNamespaces.WordprocessingML);
            _revisionData = rev;
            _currentPapx = currentPapx;
            _styleChpx = styleChpx;
            _currentIstd = UInt16.MaxValue;
        }

        public CharacterPropertiesMapping(XmlElement rPr, WordDocument doc, RevisionData rev, ParagraphPropertyExceptions currentPapx, bool styleChpx)
            : base(null)
        {
            _doc = doc;
            _nodeFactory = rPr.OwnerDocument;
            _rPr = rPr;
            _revisionData = rev;
            _currentPapx = currentPapx;
            _styleChpx = styleChpx;
            _currentIstd = UInt16.MaxValue;
        }

        public void Apply(CharacterPropertyExceptions chpx)
        {
            //convert the normal SPRMS
            convertSprms(chpx.grpprl, _rPr);

            //apend revision changes
            if (_revisionData.Type == RevisionData.RevisionType.Changed)
            {
                XmlElement rPrChange = _nodeFactory.CreateElement("w", "rPrChange", OpenXmlNamespaces.WordprocessingML);

                //date
                _revisionData.Dttm.Convert(new DateMapping(rPrChange));

                //author
                XmlAttribute author = _nodeFactory.CreateAttribute("w", "author", OpenXmlNamespaces.WordprocessingML);
                author.Value = _doc.RevisionAuthorTable.Strings[_revisionData.Isbt];
                rPrChange.Attributes.Append(author);

                //convert revision stack
                convertSprms(_revisionData.Changes, rPrChange);

                _rPr.AppendChild(rPrChange);
            }

            //write properties
            if (_writer!=null && (_rPr.ChildNodes.Count > 0 || _rPr.Attributes.Count > 0))
            {
                _rPr.WriteTo(_writer);
            }
        }

        private void convertSprms(List<SinglePropertyModifier> sprms, XmlElement parent)
        {
            XmlElement shd = _nodeFactory.CreateElement("w", "shd", OpenXmlNamespaces.WordprocessingML);
            XmlElement rFonts = _nodeFactory.CreateElement("w", "rFonts", OpenXmlNamespaces.WordprocessingML);
            XmlElement color = _nodeFactory.CreateElement("w", "color", OpenXmlNamespaces.WordprocessingML);
            XmlAttribute colorVal = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
            XmlElement lang = _nodeFactory.CreateElement("w", "lang", OpenXmlNamespaces.WordprocessingML);

            foreach (SinglePropertyModifier sprm in sprms)
            {
                switch ((int)sprm.OpCode)
                {
                    //style id 
                    case 0x4A30:
                        _currentIstd = System.BitConverter.ToUInt16(sprm.Arguments, 0);
                        appendValueElement(parent, "rStyle", StyleSheetMapping.MakeStyleId(_doc.Styles.Styles[_currentIstd]), true);
                        break;
                    
                    //Element flags
                    case 0x085A:
                        appendFlagElement(parent, sprm, "rtl", true);
                        break;
                    case 0x0835:
                        appendFlagElement(parent, sprm, "b", true);
                        break;
                    case 0x085C:
                        appendFlagElement(parent, sprm, "bCs", true);
                        break;
                    case 0x083B:
                        appendFlagElement(parent, sprm, "caps", true); ;
                        break;
                    case 0x0882:
                        appendFlagElement(parent, sprm, "cs", true);
                        break;
                    case 0x2A53:
                        appendFlagElement(parent, sprm, "dstrike", true);
                        break;
                    case 0x0858:
                        appendFlagElement(parent, sprm, "emboss", true);
                        break;
                    case 0x0854:
                        appendFlagElement(parent, sprm, "imprint", true);
                        break;
                    case 0x0836:
                        appendFlagElement(parent, sprm, "i", true);
                        break;
                    case 0x085D:
                        appendFlagElement(parent, sprm, "iCs", true);
                        break;
                    case 0x0875:
                        appendFlagElement(parent, sprm, "noProof", true);
                        break;
                    case 0x0838:
                        appendFlagElement(parent, sprm, "outline", true);
                        break;
                    case 0x0839:
                        appendFlagElement(parent, sprm, "shadow", true);
                        break;
                    case 0x083A:
                        appendFlagElement(parent, sprm, "smallCaps", true);
                        break;
                    case 0x0818:
                        appendFlagElement(parent, sprm, "specVanish", true);
                        break;
                    case 0x0837:
                        appendFlagElement(parent, sprm, "strike", true);
                        break;
                    case 0x083C:
                        appendFlagElement(parent, sprm, "vanish", true);
                        break;
                    case 0x0811:
                        appendFlagElement(parent, sprm, "webHidden", true);
                        break;
                    case 0x2A48:
                        var iss = (SuperscriptIndex)sprm.Arguments[0];
                        appendValueElement(parent, "vertAlign", iss.ToString(), true);
                        break;

                    //language
                    case 0x486D:
                    case 0x4873:
                        //latin
                        var langid = new LanguageId(System.BitConverter.ToInt16(sprm.Arguments, 0));
                        langid.Convert(new LanguageIdMapping(lang, LanguageIdMapping.LanguageType.Default));
                        break;
                    case 0x486E:
                    case 0x4874:
                        //east asia
                        langid = new LanguageId(System.BitConverter.ToInt16(sprm.Arguments, 0));
                        langid.Convert(new LanguageIdMapping(lang, LanguageIdMapping.LanguageType.EastAsian));
                        break;
                    case 0x485F:
                        //bidi
                        langid = new LanguageId(System.BitConverter.ToInt16(sprm.Arguments, 0));
                        langid.Convert(new LanguageIdMapping(lang, LanguageIdMapping.LanguageType.Complex));
                        break;
                    
                    //borders
                    case 0x6865:
                    case 0xCA72:
                        XmlNode bdr = _nodeFactory.CreateElement("w", "bdr", OpenXmlNamespaces.WordprocessingML);
                        appendBorderAttributes(new BorderCode(sprm.Arguments), bdr);
                        parent.AppendChild(bdr);
                        break;

                    //shading
                    case 0x4866:
                    case 0xCA71:
                        var desc = new ShadingDescriptor(sprm.Arguments);
                        appendShading(parent, desc);
                        break;

                    //color
                    case 0x2A42:
                    case 0x4A60:
                        colorVal.Value = ((Global.ColorIdentifier)(sprm.Arguments[0])).ToString();
                        break;
                    case 0x6870:
                        //R
                        colorVal.Value = String.Format("{0:x2}", sprm.Arguments[0]);
                        //G
                        colorVal.Value += String.Format("{0:x2}", sprm.Arguments[1]);
                        //B
                        colorVal.Value += String.Format("{0:x2}", sprm.Arguments[2]);
                        break;

                    //highlightning
                    case 0x2A0C:
                        appendValueElement(parent, "highlight", ((Global.ColorIdentifier)sprm.Arguments[0]).ToString(), true);
                        break;

                    //spacing
                    case 0x8840:
                        appendValueElement(parent, "spacing", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString(), true);
                        break;

                    //font size
                    case 0x4A43:
                        appendValueElement(parent, "sz", sprm.Arguments[0].ToString(), true);
                        break;
                    case 0x484B:
                        appendValueElement(parent, "kern", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString(), true);
                        break;
                    case 0x4A61:
                        appendValueElement(parent, "szCs", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString(), true);
                        break;
                    
                    //font family
                    case 0x4A4F:
                        XmlAttribute ascii = _nodeFactory.CreateAttribute("w", "ascii", OpenXmlNamespaces.WordprocessingML);
                        var ffn = (FontFamilyName)_doc.FontTable.Data[System.BitConverter.ToUInt16(sprm.Arguments, 0)];
                        ascii.Value = ffn.xszFtn;
                        rFonts.Attributes.Append(ascii);
                        break;
                   case 0x4A50:
                       XmlAttribute eastAsia = _nodeFactory.CreateAttribute("w", "eastAsia", OpenXmlNamespaces.WordprocessingML);
                       var ffnAsia = (FontFamilyName)_doc.FontTable.Data[System.BitConverter.ToUInt16(sprm.Arguments, 0)];
                       eastAsia.Value = ffnAsia.xszFtn;
                       rFonts.Attributes.Append(eastAsia);
                       break;
                    case 0x4A51:
                        XmlAttribute ansi = _nodeFactory.CreateAttribute("w", "hAnsi", OpenXmlNamespaces.WordprocessingML);
                        var ffnAnsi = (FontFamilyName)_doc.FontTable.Data[System.BitConverter.ToUInt16(sprm.Arguments, 0)];
                        ansi.Value = ffnAnsi.xszFtn;
                        rFonts.Attributes.Append(ansi);
                        break;
                    case (int)SinglePropertyModifier.OperationCode.sprmCIdctHint:
                        // it's complex script 
                        XmlAttribute hint = _nodeFactory.CreateAttribute("w", "hint", OpenXmlNamespaces.WordprocessingML);
                        hint.Value = "cs";
                        rFonts.Attributes.Append(hint);
                        break;
                    case (int)SinglePropertyModifier.OperationCode.sprmCFtcBi:
                        // complex script font
                        XmlAttribute cs = _nodeFactory.CreateAttribute("w", "cs", OpenXmlNamespaces.WordprocessingML);
                        var ffnCs = (FontFamilyName)_doc.FontTable.Data[System.BitConverter.ToUInt16(sprm.Arguments, 0)];
                        cs.Value = ffnCs.xszFtn;
                        rFonts.Attributes.Append(cs);
                        break;

                    //Underlining
                    case 0x2A3E:
                        appendValueElement(parent, "u", lowerFirstChar(((Global.UnderlineCode)sprm.Arguments[0]).ToString()), true);
                        break;

                    //char width
                    case 0x4852:
                        appendValueElement(parent, "w", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString(), true);
                        break;

                    //animation
                    case 0x2859:
                        appendValueElement(parent, "effect", ((Global.TextAnimation)sprm.Arguments[0]).ToString(), true);
                        break;

                    default:
                        break;
                }
            }

            //apend lang
            if (lang.Attributes.Count > 0)
            {
                parent.AppendChild(lang);
            }

            //append fonts
            if (rFonts.Attributes.Count > 0)
            {
                parent.AppendChild(rFonts);
            }

            //append color
            if (colorVal.Value != "")
            {
                color.Attributes.Append(colorVal);
                parent.AppendChild(color);
            }
        }

        /// <summary>
        /// CHPX flags are special flags because the can be 0,1,128 and 129,
        /// so this method overrides the appendFlagElement method.
        /// </summary>
        protected override void appendFlagElement(XmlElement node, SinglePropertyModifier sprm, string elementName, bool unique)
        {
            byte flag = sprm.Arguments[0];
            if(flag != 128)
            {
                XmlElement ele = _nodeFactory.CreateElement("w", elementName, OpenXmlNamespaces.WordprocessingML);
                XmlAttribute val = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);

                if (unique)
                {
                    foreach (XmlElement exEle in node.ChildNodes)
                    {
                        if (exEle.Name == ele.Name)
                        {
                            node.RemoveChild(exEle);
                            break;
                        }
                    }

                }

                if (flag == 0)
                {
                    val.Value = "false";
                    ele.Attributes.Append(val);
                    node.AppendChild(ele);
                }
                else if (flag == 1)
                {
                    //dont append attribute val
                    //no attribute means true
                    node.AppendChild(ele);
                }
                else if(flag == 129)
                {
                    //Invert the value of the style

                    //determine the style id of the current style
                    ushort styleId = 0;
                    if (_currentIstd != UInt16.MaxValue)
                    {
                        styleId = _currentIstd;
                    }
                    else if(_currentPapx != null)
                    {
                        styleId = _currentPapx.istd;
                    }

                    //this chpx is the chpx of a style, 
                    //don't use the id of the chpx or the papx, use the baseOn style
                    if (_styleChpx)
                    {
                        StyleSheetDescription thisStyle = _doc.Styles.Styles[styleId];
                        styleId = (ushort)thisStyle.istdBase;
                    }

                    //build the style hierarchy
                    if (_hierarchy == null)
                    {
                        _hierarchy = buildHierarchy(_doc.Styles, styleId);
                    }

                    //apply the toggle values to get the real value of the style
                    bool stylesVal = applyToggleHierachy(sprm); 

                    //invert it
                    if (stylesVal)
                    {
                        val.Value = "0";
                        ele.Attributes.Append(val);
                    }
                    node.AppendChild(ele);
                }
            }
        }

        private List<CharacterPropertyExceptions> buildHierarchy(StyleSheet styleSheet, ushort istdStart)
        {
            var hierarchy = new List<CharacterPropertyExceptions>();
            int istd = (int)istdStart;
            bool goOn = true;
            while (goOn)
            {
                try
                {
                    CharacterPropertyExceptions baseChpx = styleSheet.Styles[istd].chpx;
                    if (baseChpx != null)
                    {
                        hierarchy.Add(baseChpx);
                        istd = (int)styleSheet.Styles[istd].istdBase;
                    }
                    else
                    {
                        goOn = false;
                    }
                }
                catch (Exception)
                {
                    goOn = false;
                }
            }
            return hierarchy;
        }


        private bool applyToggleHierachy(SinglePropertyModifier sprm)
        {
            bool ret = false;
            foreach (CharacterPropertyExceptions ancientChpx in _hierarchy)
            {
                foreach (SinglePropertyModifier ancientSprm in ancientChpx.grpprl)
                {
                    if (ancientSprm.OpCode == sprm.OpCode)
                    {
                        byte ancient = ancientSprm.Arguments[0];
                        ret = toogleValue(ret, ancient);
                        break;
                    }
                }
            }
            return ret;
        }


        private bool toogleValue(bool currentValue, byte toggle)
        {
            if (toggle == 1)
                return true;
            else if (toggle == 129)
                //invert the current value
                if (currentValue)
                    return false;
                else
                    return true;
            else if (toggle == 128)
                //use the current value
                return currentValue;
            else
                return false;
        }


        private string lowerFirstChar(string s)
        {
            return s.Substring(0, 1).ToLower() + s.Substring(1, s.Length - 1);
        }
    }
}
