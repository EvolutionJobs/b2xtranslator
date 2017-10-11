

using System.Collections.Generic;
using b2xtranslator.PptFileFormat;
using b2xtranslator.CommonTranslatorLib;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.Tools;

namespace b2xtranslator.PresentationMLMapping
{
    class ColorSchemeMapping :
        AbstractOpenXmlMapping
    {
        protected ConversionContext _ctx;

        public ColorSchemeMapping(ConversionContext ctx, XmlWriter writer)
            : base(writer)
        {
            this._ctx = ctx;
        }

        public void Apply(List<ColorSchemeAtom> schemes)
        {
            this._writer.WriteStartElement("a", "theme", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("name", "dummyTheme");

            this._writer.WriteStartElement("a", "themeElements", OpenXmlNamespaces.DrawingML);
            var s = schemes[schemes.Count-1];
            writeScheme(s);

            //write fontScheme
            writeFontScheme();

            //write fmtScheme
            writeFmtScheme();

            this._writer.WriteEndElement(); // themeElements

            if (schemes.Count > 1)
            {
                this._writer.WriteStartElement("a", "extraClrSchemeLst", OpenXmlNamespaces.DrawingML);
                foreach (var scheme in schemes)
                {
                    this._writer.WriteStartElement("a","extraClrScheme", OpenXmlNamespaces.DrawingML);
                    writeScheme(scheme);
                    this._writer.WriteEndElement(); // extraClrScheme
                }
                this._writer.WriteEndElement(); // extraClrSchemeLst
            }


            this._writer.WriteEndElement(); // theme
            
        }

        private void writeFmtScheme()
        {
            //TODO: real implementation instead of default 
            var xmlDoc = Utils.GetDefaultDocument("theme");
            var l = xmlDoc.GetElementsByTagName("fmtScheme", OpenXmlNamespaces.DrawingML);

            l.Item(0).WriteTo(this._writer);

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

            this._writer.WriteStartElement("a", "fontScheme", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("name", "dummyFontScheme");

            this._writer.WriteStartElement("a", "majorFont", OpenXmlNamespaces.DrawingML);
            this._writer.WriteStartElement("a", "latin", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("typeface", "Arial Black");
            this._writer.WriteEndElement(); // latin
            this._writer.WriteStartElement("a", "ea", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("typeface", "");
            this._writer.WriteEndElement(); // ea
            this._writer.WriteStartElement("a", "cs", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("typeface", "");
            this._writer.WriteEndElement(); // cs
            this._writer.WriteEndElement(); // majorFont

            this._writer.WriteStartElement("a", "minorFont", OpenXmlNamespaces.DrawingML);
            this._writer.WriteStartElement("a", "latin", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("typeface", "Arial");
            this._writer.WriteEndElement(); // latin
            this._writer.WriteStartElement("a", "ea", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("typeface", "");
            this._writer.WriteEndElement(); // ea
            this._writer.WriteStartElement("a", "cs", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("typeface", "");
            this._writer.WriteEndElement(); // cs
            this._writer.WriteEndElement(); // minorFont

            this._writer.WriteEndElement(); // fontScheme
        }

        private void writeScheme(ColorSchemeAtom scheme)
        {
            //TODO: verify the mappings; at leat accent4 - 6 are wrong

            this._writer.WriteStartElement("a", "clrScheme", OpenXmlNamespaces.DrawingML);

            this._writer.WriteAttributeString("name", "dummyScheme");

            this._writer.WriteStartElement("a", "dk1", OpenXmlNamespaces.DrawingML);
            this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("val", new RGBColor(scheme.TextAndLines, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            this._writer.WriteEndElement(); // srgbClr
            this._writer.WriteEndElement(); // dk1

            this._writer.WriteStartElement("a", "lt1", OpenXmlNamespaces.DrawingML);
            this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("val", new RGBColor(scheme.Background, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            this._writer.WriteEndElement(); // srgbClr
            this._writer.WriteEndElement(); // lt1

            this._writer.WriteStartElement("a", "dk2", OpenXmlNamespaces.DrawingML);
            this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("val", new RGBColor(scheme.TitleText, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            this._writer.WriteEndElement(); // srgbClr
            this._writer.WriteEndElement(); // dk2

            this._writer.WriteStartElement("a", "lt2", OpenXmlNamespaces.DrawingML);
            this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("val", new RGBColor(scheme.Shadows, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            this._writer.WriteEndElement(); // srgbClr
            this._writer.WriteEndElement(); // lt2

            this._writer.WriteStartElement("a", "accent1", OpenXmlNamespaces.DrawingML);
            this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("val", new RGBColor(scheme.Fills, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            this._writer.WriteEndElement(); // srgbClr
            this._writer.WriteEndElement(); // accent1

            this._writer.WriteStartElement("a", "accent2", OpenXmlNamespaces.DrawingML);
            this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("val", new RGBColor(scheme.Accent, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            this._writer.WriteEndElement(); // srgbClr
            this._writer.WriteEndElement(); // accent2

            this._writer.WriteStartElement("a", "accent3", OpenXmlNamespaces.DrawingML);
            this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("val", new RGBColor(scheme.Background, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            this._writer.WriteEndElement(); // srgbClr
            this._writer.WriteEndElement(); // accent3

            this._writer.WriteStartElement("a", "accent4", OpenXmlNamespaces.DrawingML);
            this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("val", new RGBColor(scheme.Accent, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            this._writer.WriteEndElement(); // srgbClr
            this._writer.WriteEndElement(); // accent4

            this._writer.WriteStartElement("a", "accent5", OpenXmlNamespaces.DrawingML);
            this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("val", new RGBColor(scheme.Accent, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            this._writer.WriteEndElement(); // srgbClr
            this._writer.WriteEndElement(); // accent5

            this._writer.WriteStartElement("a", "accent6", OpenXmlNamespaces.DrawingML);
            this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("val", new RGBColor(scheme.Accent, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            this._writer.WriteEndElement(); // srgbClr
            this._writer.WriteEndElement(); // accent6

            this._writer.WriteStartElement("a", "hlink", OpenXmlNamespaces.DrawingML);
            this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("val", new RGBColor(scheme.AccentAndHyperlink, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            this._writer.WriteEndElement(); // srgbClr
            this._writer.WriteEndElement(); // hlink

            this._writer.WriteStartElement("a", "folHlink", OpenXmlNamespaces.DrawingML);
            this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("val", new RGBColor(scheme.AccentAndFollowedHyperlink, RGBColor.ByteOrder.RedFirst).SixDigitHexCode);
            this._writer.WriteEndElement(); // srgbClr
            this._writer.WriteEndElement(); // folHlink

            this._writer.WriteEndElement(); // clrScheme


        }
    }
}
