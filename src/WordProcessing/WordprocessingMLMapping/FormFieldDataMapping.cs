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
            this._writer.WriteStartElement("w", "ffData", OpenXmlNamespaces.WordprocessingML);

            //name
            this._writer.WriteStartElement("w", "name", OpenXmlNamespaces.WordprocessingML);
            this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.xstzName);
            this._writer.WriteEndElement();

            //calcOnExit
            this._writer.WriteStartElement("w", "calcOnExit", OpenXmlNamespaces.WordprocessingML);
            if (ffd.fRecalc)
                this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "1");
            else
                this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "0");
            this._writer.WriteEndElement();

            //entry macro
            if (ffd.xstzEntryMcr != null && ffd.xstzEntryMcr.Length > 0)
            {
                this._writer.WriteStartElement("w", "entryMacro", OpenXmlNamespaces.WordprocessingML);
                this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.xstzEntryMcr);
                this._writer.WriteEndElement();
            }

            //exit macro
            if (ffd.xstzExitMcr != null && ffd.xstzExitMcr.Length > 0)
            {
                this._writer.WriteStartElement("w", "exitMacro", OpenXmlNamespaces.WordprocessingML);
                this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.xstzExitMcr);
                this._writer.WriteEndElement();
            }

            //help text
            if (ffd.xstzHelpText != null && ffd.xstzHelpText.Length > 0)
            {
                this._writer.WriteStartElement("w", "helpText", OpenXmlNamespaces.WordprocessingML);
                this._writer.WriteAttributeString("w", "type", OpenXmlNamespaces.WordprocessingML, "text");
                this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.xstzHelpText);
                this._writer.WriteEndElement();
            }

            //status text
            if (ffd.xstzStatText != null && ffd.xstzStatText.Length > 0)
            {
                this._writer.WriteStartElement("w", "statusText", OpenXmlNamespaces.WordprocessingML);
                this._writer.WriteAttributeString("w", "type", OpenXmlNamespaces.WordprocessingML, "text");
                this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.xstzStatText);
                this._writer.WriteEndElement();
            }

            //start custom properties
            switch (ffd.iType)
            {
                case FormFieldData.FormFieldType.iTypeText:
                    this._writer.WriteStartElement("w", "textInput", OpenXmlNamespaces.WordprocessingML);

                    //type
                    if (ffd.iTypeTxt != FormFieldData.TextboxType.regular)
                    {
                        this._writer.WriteStartElement("w", "type", OpenXmlNamespaces.WordprocessingML);
                        this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.iTypeTxt.ToString());
                        this._writer.WriteEndElement();
                    }

                    //length
                    if (ffd.cch > 0)
                    {
                        this._writer.WriteStartElement("w", "maxLength", OpenXmlNamespaces.WordprocessingML);
                        this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.cch.ToString());
                        this._writer.WriteEndElement();
                    }
                    
                    //textformat
                    if (ffd.xstzTextFormat != null && ffd.xstzTextFormat.Length > 0)
                    {
                        this._writer.WriteStartElement("w", "format", OpenXmlNamespaces.WordprocessingML);
                        this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.xstzTextFormat);
                        this._writer.WriteEndElement();
                    }

                    //default text
                    if (ffd.xstzTextDef != null && ffd.xstzTextDef.Length > 0)
                    {
                        this._writer.WriteStartElement("w", "default", OpenXmlNamespaces.WordprocessingML);
                        this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.xstzTextDef);
                        this._writer.WriteEndElement();
                    }

                    break;
                case FormFieldData.FormFieldType.iTypeChck:
                    this._writer.WriteStartElement("w", "checkBox", OpenXmlNamespaces.WordprocessingML);

                    //checked <w:checked w:val="0"/>
                    if (ffd.iRes != UNDEFINED_RESULT)
                    {
                        this._writer.WriteStartElement("w", "checked", OpenXmlNamespaces.WordprocessingML);
                        this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.iRes.ToString());
                        this._writer.WriteEndElement();
                    }

                    //size 
                    if (ffd.hps >= 2 && ffd.hps <= 3168)
                    {
                        this._writer.WriteStartElement("w", "size", OpenXmlNamespaces.WordprocessingML);
                        this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.hps.ToString());
                        this._writer.WriteEndElement();
                    }
                    else
                    {
                        this._writer.WriteStartElement("w", "sizeAuto", OpenXmlNamespaces.WordprocessingML);
                        this._writer.WriteEndElement();
                    }

                    //default setting
                    this._writer.WriteStartElement("w", "default", OpenXmlNamespaces.WordprocessingML);
                    this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.wDef.ToString());
                    this._writer.WriteEndElement();
                    
                    break;
                case FormFieldData.FormFieldType.iTypeDrop:
                    this._writer.WriteStartElement("w", "ddList", OpenXmlNamespaces.WordprocessingML);

                    //selected item
                    if (ffd.iRes != UNDEFINED_RESULT)
                    {
                        this._writer.WriteStartElement("w", "result", OpenXmlNamespaces.WordprocessingML);
                        this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.iRes.ToString());
                        this._writer.WriteEndElement();
                    }

                    //default

                    //_entries

                    break;
                default:
                    this._writer.WriteStartElement("w", "textInput", OpenXmlNamespaces.WordprocessingML);

                    break;
            }

            //close custom properties
            this._writer.WriteEndElement();

            //close ffData
            this._writer.WriteEndElement();
        }
    }
}
