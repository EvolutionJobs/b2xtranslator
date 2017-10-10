/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
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

using System.Xml;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.StyleData;


namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{

    public class SSTMapping : AbstractOpenXmlMapping,
          IMapping<SSTData>
    {
        ExcelContext xlsContext;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="xlsContext">The excel context object</param>
        public SSTMapping(ExcelContext xlsContext)
            : base(XmlWriter.Create(xlsContext.SpreadDoc.WorkbookPart.AddSharedStringPart().GetStream(), xlsContext.WriterSettings))
        {
            this.xlsContext = xlsContext;
        }

        /// <summary>
        /// The overload apply method 
        /// Creates the sharedstring xml document 
        /// </summary>
        /// <param name="SSTData">SharedStringData Object</param>
        public void Apply(SSTData sstData)
        {
            this._writer.WriteStartDocument();
            this._writer.WriteStartElement("sst", OpenXmlNamespaces.SpreadsheetML);
            // count="x" uniqueCount="y" 
            this._writer.WriteAttributeString("count", sstData.cstTotal.ToString());
            this._writer.WriteAttributeString("uniqueCount", sstData.cstUnique.ToString());



            int count = 0;
            // create the string _entries 
            foreach (var var in sstData.StringList)
            {
                count++;
                var list = sstData.getFormatingRuns(count);

                this._writer.WriteStartElement("si");

                if (list.Count == 0)
                {
                    // if there is no formatting, there is no run, write only the text
                    writeTextNode(this._writer, var);
                }
                else
                {
                    // if there is no formatting, there is no run, write only the text

                    // first text 
                    if (list[0].CharNumber != 0)
                    {
                        // no formating for the first letters 
                        this._writer.WriteStartElement("r");
                        writeTextNode(this._writer, var.Substring(0, list[0].CharNumber));
                        this._writer.WriteEndElement();
                    }

                    FontData fd;
                    for (int i = 0; i <= list.Count - 2; i++)
                    {
                        this._writer.WriteStartElement("r");

                        fd = this.xlsContext.XlsDoc.WorkBookData.styleData.FontDataList[list[i].FontRecord];
                        StyleMappingHelper.addFontElement(this._writer, fd, FontElementType.String);

                        writeTextNode(this._writer, var.Substring(list[i].CharNumber, list[i + 1].CharNumber - list[i].CharNumber));
                        this._writer.WriteEndElement();
                    }
                    this._writer.WriteStartElement("r");

                    fd = this.xlsContext.XlsDoc.WorkBookData.styleData.FontDataList[list[list.Count - 1].FontRecord];
                    StyleMappingHelper.addFontElement(this._writer, fd, FontElementType.String);

                    writeTextNode(this._writer, var.Substring(list[list.Count - 1].CharNumber));
                    this._writer.WriteEndElement();
                }

                this._writer.WriteEndElement(); // end si

            }

            // close tags 
            this._writer.WriteEndElement();
            this._writer.WriteEndDocument();

            // close writer 
            this._writer.Flush();
        }


        private void writeTextNode(XmlWriter writer, string text)
        {
            writer.WriteStartElement("t");
            if ( text.StartsWith(" ") || text.EndsWith(" ") ||
                text.StartsWith("\n") || text.EndsWith("\n") ||
                text.StartsWith("\r") || text.EndsWith("\r") ) 
            {
                writer.WriteAttributeString("xml", "space", "", "preserve");
            }
            writer.WriteString(text);
            writer.WriteEndElement();
        }

    }

}
