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

using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class FormFieldDataMapping :
        AbstractOpenXmlMapping,
        IMapping<FormFieldData>
    {
        private const int UNDEFINED_RESULT = 25;

        public FormFieldDataMapping(XmlWriter writer)
            : base(writer)
        {
        }

        public void Apply(FormFieldData ffd)
        {
            _writer.WriteStartElement("w", "ffData", OpenXmlNamespaces.WordprocessingML);

            //name
            _writer.WriteStartElement("w", "name", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.xstzName);
            _writer.WriteEndElement();

            //calcOnExit
            _writer.WriteStartElement("w", "calcOnExit", OpenXmlNamespaces.WordprocessingML);
            if (ffd.fRecalc)
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "1");
            else
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "0");
            _writer.WriteEndElement();

            //entry macro
            if (ffd.xstzEntryMcr != null && ffd.xstzEntryMcr.Length > 0)
            {
                _writer.WriteStartElement("w", "entryMacro", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.xstzEntryMcr);
                _writer.WriteEndElement();
            }

            //exit macro
            if (ffd.xstzExitMcr != null && ffd.xstzExitMcr.Length > 0)
            {
                _writer.WriteStartElement("w", "exitMacro", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.xstzExitMcr);
                _writer.WriteEndElement();
            }

            //help text
            if (ffd.xstzHelpText != null && ffd.xstzHelpText.Length > 0)
            {
                _writer.WriteStartElement("w", "helpText", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "type", OpenXmlNamespaces.WordprocessingML, "text");
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.xstzHelpText);
                _writer.WriteEndElement();
            }

            //status text
            if (ffd.xstzStatText != null && ffd.xstzStatText.Length > 0)
            {
                _writer.WriteStartElement("w", "statusText", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "type", OpenXmlNamespaces.WordprocessingML, "text");
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.xstzStatText);
                _writer.WriteEndElement();
            }

            //start custom properties
            switch (ffd.iType)
            {
                case FormFieldData.FormFieldType.iTypeText:
                    _writer.WriteStartElement("w", "textInput", OpenXmlNamespaces.WordprocessingML);

                    //type
                    if (ffd.iTypeTxt != FormFieldData.TextboxType.regular)
                    {
                        _writer.WriteStartElement("w", "type", OpenXmlNamespaces.WordprocessingML);
                        _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.iTypeTxt.ToString());
                        _writer.WriteEndElement();
                    }

                    //length
                    if (ffd.cch > 0)
                    {
                        _writer.WriteStartElement("w", "maxLength", OpenXmlNamespaces.WordprocessingML);
                        _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.cch.ToString());
                        _writer.WriteEndElement();
                    }
                    
                    //textformat
                    if (ffd.xstzTextFormat != null && ffd.xstzTextFormat.Length > 0)
                    {
                        _writer.WriteStartElement("w", "format", OpenXmlNamespaces.WordprocessingML);
                        _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.xstzTextFormat);
                        _writer.WriteEndElement();
                    }

                    //default text
                    if (ffd.xstzTextDef != null && ffd.xstzTextDef.Length > 0)
                    {
                        _writer.WriteStartElement("w", "default", OpenXmlNamespaces.WordprocessingML);
                        _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.xstzTextDef);
                        _writer.WriteEndElement();
                    }

                    break;
                case FormFieldData.FormFieldType.iTypeChck:
                    _writer.WriteStartElement("w", "checkBox", OpenXmlNamespaces.WordprocessingML);

                    //checked <w:checked w:val="0"/>
                    if (ffd.iRes != UNDEFINED_RESULT)
                    {
                        _writer.WriteStartElement("w", "checked", OpenXmlNamespaces.WordprocessingML);
                        _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.iRes.ToString());
                        _writer.WriteEndElement();
                    }

                    //size 
                    if (ffd.hps >= 2 && ffd.hps <= 3168)
                    {
                        _writer.WriteStartElement("w", "size", OpenXmlNamespaces.WordprocessingML);
                        _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.hps.ToString());
                        _writer.WriteEndElement();
                    }
                    else
                    {
                        _writer.WriteStartElement("w", "sizeAuto", OpenXmlNamespaces.WordprocessingML);
                        _writer.WriteEndElement();
                    }

                    //default setting
                    _writer.WriteStartElement("w", "default", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.wDef.ToString());
                    _writer.WriteEndElement();
                    
                    break;
                case FormFieldData.FormFieldType.iTypeDrop:
                    _writer.WriteStartElement("w", "ddList", OpenXmlNamespaces.WordprocessingML);

                    //selected item
                    if (ffd.iRes != UNDEFINED_RESULT)
                    {
                        _writer.WriteStartElement("w", "result", OpenXmlNamespaces.WordprocessingML);
                        _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.iRes.ToString());
                        _writer.WriteEndElement();
                    }

                    //default

                    //_entries

                    break;
                default:
                    _writer.WriteStartElement("w", "textInput", OpenXmlNamespaces.WordprocessingML);

                    break;
            }

            //close custom properties
            _writer.WriteEndElement();

            //close ffData
            _writer.WriteEndElement();
        }
    }
}
