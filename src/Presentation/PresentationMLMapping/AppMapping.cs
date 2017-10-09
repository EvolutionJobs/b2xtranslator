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
using DIaLOGIKa.b2xtranslator.PptFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.PresentationML;
using DIaLOGIKa.b2xtranslator.Tools;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    public class AppMapping : AbstractOpenXmlMapping,
          IMapping<IVisitable>
    {

        public AppMapping(AppPropertiesPart appPart, XmlWriterSettings xws)
            : base(XmlWriter.Create(appPart.GetStream(), xws))
        {
        }

        public void Apply(IVisitable x)
        {
            // Start the document
            _writer.WriteStartDocument();
            _writer.WriteStartElement("Properties", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
                               
            // Force declaration of these namespaces at document start
            //_writer.WriteAttributeString("xmlns", null, null, OpenXmlRelationshipTypes.ExtendedProperties);
            _writer.WriteAttributeString("xmlns", "vt", null, OpenXmlNamespaces.docPropsVTypes);

            //TotalTime
            _writer.WriteElementString("TotalTime", "0");
            //Words
            _writer.WriteElementString("Words", "0");
            //PresentationFormat
            _writer.WriteElementString("PresentationFormat", "Custom");
            //Paragraphs
            _writer.WriteElementString("Paragraphs", "0");
            //Slides
            _writer.WriteElementString("Slides", "0");
            //Notes
            _writer.WriteElementString("Notes", "0");
            //HiddenSlides
            _writer.WriteElementString("HiddenSlides", "0");
            //MMClips
            _writer.WriteElementString("MMClips", "0");
            //ScaleCrop
            _writer.WriteElementString("ScaleCrop", "false");
            //HeadingPairs
            //TitlesOfParts           

            //LinksUpToDate
            _writer.WriteElementString("LinksUpToDate", "false");
            //SharedDoc
            _writer.WriteElementString("SharedDoc", "false");
            //HyperlinksChanged
            _writer.WriteElementString("HyperlinksChanged", "false");
            //AppVersion
            _writer.WriteElementString("AppVersion", "12.0000");
            
            // End the document
            _writer.WriteEndElement(); //Properties
            _writer.WriteEndDocument();

            _writer.Flush();
        }        
    }
}
