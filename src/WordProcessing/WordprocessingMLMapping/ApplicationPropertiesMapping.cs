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
    public class ApplicationPropertiesMapping : AbstractOpenXmlMapping,
          IMapping<DocumentProperties>
    {

        public ApplicationPropertiesMapping(AppPropertiesPart appPart, XmlWriterSettings xws)
            : base(XmlWriter.Create(appPart.GetStream(), xws))
        {
        }

        public void Apply(DocumentProperties dop)
        {
            //start Properties
            _writer.WriteStartElement("w", "Properties", OpenXmlNamespaces.WordprocessingML);

            //Application
            //AppVersion
            //Company
            //DigSig
            //DocSecurity
            //HeadingPairs
            //HiddenSlides
            //HLinks
            //HyperlinkBase
            //HyperlinksChanged
            //LinksUpToDate
            //Manager
            //MMClips
            //Notes
            //PresentationFormat
            //ScaleCrop
            //SharedDoc
            //Slides
            //Template
            //TitlesOfParts
            //TotalTime

            //WordCount statistics

            //CharactersWithSpaces
            _writer.WriteStartElement("CharactersWithSpaces");
            _writer.WriteString(dop.cChWS.ToString());
            _writer.WriteEndElement();

            //Characters
            _writer.WriteStartElement("Characters");
            _writer.WriteString(dop.cCh.ToString());
            _writer.WriteEndElement();

            //Lines
            _writer.WriteStartElement("Lines");
            _writer.WriteString(dop.cLines.ToString());
            _writer.WriteEndElement();

            //Pages
            _writer.WriteStartElement("Pages");
            _writer.WriteString(dop.cPg.ToString());
            _writer.WriteEndElement();

            //Paragraphs
            _writer.WriteStartElement("Paragraphs");
            _writer.WriteString(dop.cParas.ToString());
            _writer.WriteEndElement();

            //Words
            _writer.WriteStartElement("Words");
            _writer.WriteString(dop.cWords.ToString());
            _writer.WriteEndElement();

            //end Properties
            _writer.WriteEndElement();
        }
    }
}
