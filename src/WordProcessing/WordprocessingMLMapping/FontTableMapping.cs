using System;
using b2xtranslator.CommonTranslatorLib;
using System.Xml;
using b2xtranslator.DocFileFormat;
using b2xtranslator.OpenXmlLib;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class FontTableMapping : AbstractOpenXmlMapping,
        IMapping<StringTable>
    {
        protected enum FontFamily
        {
            auto,
            decorative,
            modern,
            roman,
            script,
            swiss
        }

        public FontTableMapping(ConversionContext ctx, OpenXmlPart targetPart)
            : base(XmlWriter.Create(targetPart.GetStream(), ctx.WriterSettings))
        {
        }

        public void Apply(StringTable table)
        {
            this._writer.WriteStartElement("w", "fonts", OpenXmlNamespaces.WordprocessingML);

            foreach (FontFamilyName font in table.Data)
            {
                this._writer.WriteStartElement("w", "font", OpenXmlNamespaces.WordprocessingML);
                this._writer.WriteAttributeString("w", "name", OpenXmlNamespaces.WordprocessingML, font.xszFtn);

                //alternative name
                if (font.xszAlt!= null && font.xszAlt.Length > 0)
                {
                    this._writer.WriteStartElement("w", "altName", OpenXmlNamespaces.WordprocessingML);
                    this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, font.xszAlt);
                    this._writer.WriteEndElement();
                }

                //charset
                this._writer.WriteStartElement("w", "charset", OpenXmlNamespaces.WordprocessingML);
                this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, string.Format("{0:x2}", font.chs));
                this._writer.WriteEndElement();

                //font family
                this._writer.WriteStartElement("w", "family", OpenXmlNamespaces.WordprocessingML);
                this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ((FontFamily)font.ff).ToString());
                this._writer.WriteEndElement();

                //panose
                this._writer.WriteStartElement("w", "panose1", OpenXmlNamespaces.WordprocessingML);
                this._writer.WriteStartAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                foreach (byte b in font.panose)
                {
                    this._writer.WriteString(string.Format("{0:x2}", b));
                }
                this._writer.WriteEndAttribute();
                this._writer.WriteEndElement();

                //pitch
                this._writer.WriteStartElement("w", "pitch", OpenXmlNamespaces.WordprocessingML);
                this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, font.prq.ToString());
                this._writer.WriteEndElement();

                //truetype
                if (!font.fTrueType)
                {
                    this._writer.WriteStartElement("w", "notTrueType", OpenXmlNamespaces.WordprocessingML);
                    this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "true");
                    this._writer.WriteEndElement();
                }

                //font signature
                this._writer.WriteStartElement("w", "sig", OpenXmlNamespaces.WordprocessingML);
                this._writer.WriteAttributeString("w", "usb0", OpenXmlNamespaces.WordprocessingML, string.Format("{0:x8}", font.fs.UnicodeSubsetBitfield0));
                this._writer.WriteAttributeString("w", "usb1", OpenXmlNamespaces.WordprocessingML, string.Format("{0:x8}", font.fs.UnicodeSubsetBitfield1));
                this._writer.WriteAttributeString("w", "usb2", OpenXmlNamespaces.WordprocessingML, string.Format("{0:x8}", font.fs.UnicodeSubsetBitfield2));
                this._writer.WriteAttributeString("w", "usb3", OpenXmlNamespaces.WordprocessingML, string.Format("{0:x8}", font.fs.UnicodeSubsetBitfield3));
                this._writer.WriteAttributeString("w", "csb0", OpenXmlNamespaces.WordprocessingML, string.Format("{0:x8}", font.fs.CodePageBitfield0));
                this._writer.WriteAttributeString("w", "csb1", OpenXmlNamespaces.WordprocessingML, string.Format("{0:x8}", font.fs.CodePageBitfield1));
                this._writer.WriteEndElement();

                this._writer.WriteEndElement();
            }

            this._writer.WriteEndElement();

            this._writer.Flush();
        }
    }
}
