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

using System.Collections.Generic;
using DIaLOGIKa.b2xtranslator.PptFileFormat;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    class ColorSchemeMapping :
        AbstractOpenXmlMapping
    {
        protected ConversionContext _ctx;

        public ColorSchemeMapping(ConversionContext ctx, XmlWriter writer)
            : base(writer)
        {
            _ctx = ctx;
        }

        public void Apply(List<ColorSchemeAtom> schemes)
        {
            _writer.WriteStartElement("a", "theme", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("name", "dummyTheme");

            _writer.WriteStartElement("a", "themeElements", OpenXmlNamespaces.DrawingML);
            var s = schemes[schemes.Count-1];
            writeScheme(s);

            //write fontScheme
            writeFontScheme();

            //write fmtScheme
            writeFmtScheme();

            _writer.WriteEndElement(); // themeElements

            if (schemes.Count > 1)
            {
                _writer.WriteStartElement("a", "extraClrSchemeLst", OpenXmlNamespaces.DrawingML);
                foreach (var scheme in schemes)
                {
                    _writer.WriteStartElement("a","extraClrScheme", OpenXmlNamespaces.DrawingML);
                    writeScheme(scheme);
                    _writer.WriteEndElement(); // extraClrScheme
                }
                _writer.WriteEndElement(); // extraClrSchemeLst
            }


            _writer.WriteEndElement(); // theme
            
        }

        private void writeFmtScheme()
        {
            //TODO: real implementation instead of default 
            var xmlDoc = Utils.GetDefaultDocument("theme");
            var l = xmlDoc.GetElementsByTagName("fmtScheme", OpenXmlNamespaces.DrawingML);

            l.Item(0).WriteTo(_writer);

            //_writer.WriteStartElement("a", "fmtScheme", OpenXmlNamespaces.DrawingML);
            //_writer.WriteAttributeString("name", "dummyFmtScheme");

            ////fillStyleList

            ////lnStyleLst

            ////effectStyleLst

            ////bgFillStyleLst

            //_writer.WriteEndElement(); // fmtScheme
        }

        private void writeFontScheme()
        {
            //TODO: real implementation instead of default 

            _writer.WriteStartElement("a", "fontScheme", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("name", "dummyFontScheme");

            _writer.WriteStartElement("a", "majorFont", OpenXmlNamespaces.DrawingML);
            _writer.WriteStartElement("a", "latin", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("typeface", "Arial Black");
            _writer.WriteEndElement(); // latin
            _writer.WriteStartElement("a", "ea", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("typeface", "");
            _writer.WriteEndElement(); // ea
            _writer.WriteStartElement("a", "cs", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("typeface", "");
            _writer.WriteEndElement(); // cs
            _writer.WriteEndElement(); // majorFont

            _writer.WriteStartElement("a", "minorFont", OpenXmlNamespaces.DrawingML);
            _writer.WriteStartElement("a", "latin", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("typeface", "Arial");
            _writer.WriteEndElement(); // latin
            _writer.WriteStartElement("a", "ea", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("typeface", "");
            _writer.WriteEndElement(); // ea
            _writer.WriteStartElement("a", "cs", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("typeface", "");
            _writer.WriteEndElement(); // cs
            _writer.WriteEndElement(); // minorFont

            _writer.WriteEndElement(); // fontScheme
        }

        private void writeScheme(ColorSchemeAtom scheme)
        {
            //TODO: verify the mappings; at leat accent4 - 6 are wrong

            _writer.WriteStartElement("a", "clrScheme", OpenXmlNamespaces.DrawingML);
            
            _writer.WriteAttributeString("name", "dummyScheme");

            _writer.WriteStartElement("a", "dk1", OpenXmlNamespaces.DrawingML);
            _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("val", new RGBColor(scheme.TextAndLines, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            _writer.WriteEndElement(); // srgbClr
            _writer.WriteEndElement(); // dk1

            _writer.WriteStartElement("a", "lt1", OpenXmlNamespaces.DrawingML);
            _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("val", new RGBColor(scheme.Background, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            _writer.WriteEndElement(); // srgbClr
            _writer.WriteEndElement(); // lt1

            _writer.WriteStartElement("a", "dk2", OpenXmlNamespaces.DrawingML);
            _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("val", new RGBColor(scheme.TitleText, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            _writer.WriteEndElement(); // srgbClr
            _writer.WriteEndElement(); // dk2

            _writer.WriteStartElement("a", "lt2", OpenXmlNamespaces.DrawingML);
            _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("val", new RGBColor(scheme.Shadows, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            _writer.WriteEndElement(); // srgbClr
            _writer.WriteEndElement(); // lt2

            _writer.WriteStartElement("a", "accent1", OpenXmlNamespaces.DrawingML);
            _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("val", new RGBColor(scheme.Fills, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            _writer.WriteEndElement(); // srgbClr
            _writer.WriteEndElement(); // accent1

            _writer.WriteStartElement("a", "accent2", OpenXmlNamespaces.DrawingML);
            _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("val", new RGBColor(scheme.Accent, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            _writer.WriteEndElement(); // srgbClr
            _writer.WriteEndElement(); // accent2

            _writer.WriteStartElement("a", "accent3", OpenXmlNamespaces.DrawingML);
            _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("val", new RGBColor(scheme.Background, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            _writer.WriteEndElement(); // srgbClr
            _writer.WriteEndElement(); // accent3

            _writer.WriteStartElement("a", "accent4", OpenXmlNamespaces.DrawingML);
            _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("val", new RGBColor(scheme.Accent, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            _writer.WriteEndElement(); // srgbClr
            _writer.WriteEndElement(); // accent4

            _writer.WriteStartElement("a", "accent5", OpenXmlNamespaces.DrawingML);
            _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("val", new RGBColor(scheme.Accent, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            _writer.WriteEndElement(); // srgbClr
            _writer.WriteEndElement(); // accent5

            _writer.WriteStartElement("a", "accent6", OpenXmlNamespaces.DrawingML);
            _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("val", new RGBColor(scheme.Accent, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            _writer.WriteEndElement(); // srgbClr
            _writer.WriteEndElement(); // accent6

            _writer.WriteStartElement("a", "hlink", OpenXmlNamespaces.DrawingML);
            _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("val", new RGBColor(scheme.AccentAndHyperlink, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            _writer.WriteEndElement(); // srgbClr
            _writer.WriteEndElement(); // hlink

            _writer.WriteStartElement("a", "folHlink", OpenXmlNamespaces.DrawingML);
            _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("val", new RGBColor(scheme.AccentAndFollowedHyperlink, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            _writer.WriteEndElement(); // srgbClr
            _writer.WriteEndElement(); // folHlink

            _writer.WriteEndElement(); // clrScheme


        }
    }
}
