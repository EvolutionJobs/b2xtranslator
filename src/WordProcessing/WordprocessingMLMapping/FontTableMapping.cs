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
using DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
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
            _writer.WriteStartElement("w", "fonts", OpenXmlNamespaces.WordprocessingML);

            foreach (FontFamilyName font in table.Data)
            {
                _writer.WriteStartElement("w", "font", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "name", OpenXmlNamespaces.WordprocessingML, font.xszFtn);

                //alternative name
                if (font.xszAlt!= null && font.xszAlt.Length > 0)
                {
                    _writer.WriteStartElement("w", "altName", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, font.xszAlt);
                    _writer.WriteEndElement();
                }

                //charset
                _writer.WriteStartElement("w", "charset", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, String.Format("{0:x2}", font.chs));
                _writer.WriteEndElement();

                //font family
                _writer.WriteStartElement("w", "family", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ((FontFamily)font.ff).ToString());
                _writer.WriteEndElement();

                //panose
                _writer.WriteStartElement("w", "panose1", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteStartAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                foreach (byte b in font.panose)
                {
                    _writer.WriteString(String.Format("{0:x2}", b));
                }
                _writer.WriteEndAttribute();
                _writer.WriteEndElement();

                //pitch
                _writer.WriteStartElement("w", "pitch", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, font.prq.ToString());
                _writer.WriteEndElement();

                //truetype
                if (!font.fTrueType)
                {
                    _writer.WriteStartElement("w", "notTrueType", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "true");
                    _writer.WriteEndElement();
                }

                //font signature
                _writer.WriteStartElement("w", "sig", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "usb0", OpenXmlNamespaces.WordprocessingML, String.Format("{0:x8}", font.fs.UnicodeSubsetBitfield0));
                _writer.WriteAttributeString("w", "usb1", OpenXmlNamespaces.WordprocessingML, String.Format("{0:x8}", font.fs.UnicodeSubsetBitfield1));
                _writer.WriteAttributeString("w", "usb2", OpenXmlNamespaces.WordprocessingML, String.Format("{0:x8}", font.fs.UnicodeSubsetBitfield2));
                _writer.WriteAttributeString("w", "usb3", OpenXmlNamespaces.WordprocessingML, String.Format("{0:x8}", font.fs.UnicodeSubsetBitfield3));
                _writer.WriteAttributeString("w", "csb0", OpenXmlNamespaces.WordprocessingML, String.Format("{0:x8}", font.fs.CodePageBitfield0));
                _writer.WriteAttributeString("w", "csb1", OpenXmlNamespaces.WordprocessingML, String.Format("{0:x8}", font.fs.CodePageBitfield1));
                _writer.WriteEndElement();

                _writer.WriteEndElement();
            }

            _writer.WriteEndElement();

            _writer.Flush();
        }
    }
}
