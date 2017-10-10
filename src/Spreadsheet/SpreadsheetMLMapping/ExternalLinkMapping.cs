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

using System;
using System.Globalization;
using System.Xml;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class ExternalLinkMapping : AbstractOpenXmlMapping,
          IMapping<SupBookData>
    {
        ExcelContext xlsContext;


                /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="xlsContext">The excel context object</param>
        public ExternalLinkMapping(ExcelContext xlsContext)
            : base(XmlWriter.Create(xlsContext.SpreadDoc.WorkbookPart.AddExternalLinkPart().GetStream(), xlsContext.WriterSettings))
        {
            this.xlsContext = xlsContext;
            
        }

        /// <summary>
        /// The overload apply method 
        /// Creates the Workbook xml document 
        /// </summary>
        /// <param name="bsd">WorkSheetData</param>
        public void Apply(SupBookData sbd)
        {
            var uri = new Uri(sbd.VirtPath, UriKind.RelativeOrAbsolute);
            ExternalRelationship er = this.xlsContext.SpreadDoc.WorkbookPart.GetExternalLinkPart().AddExternalRelationship(OpenXmlRelationshipTypes.ExternalLinkPath, uri);

            
            
            _writer.WriteStartDocument();
            _writer.WriteStartElement("externalLink", OpenXmlNamespaces.SpreadsheetML);

            _writer.WriteStartElement("externalBook");
            _writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, er.Id.ToString());

            _writer.WriteStartElement("sheetNames");
            foreach (var var in sbd.RGST)
            {
                _writer.WriteStartElement("sheetName");
                _writer.WriteAttributeString("val", var);
                _writer.WriteEndElement(); 
            }
            _writer.WriteEndElement();

            // checks if some externNames exist
            if (sbd.ExternNames.Count > 0)
            {
                _writer.WriteStartElement("definedNames");
                foreach (var var in sbd.ExternNames)
                {
                    _writer.WriteStartElement("definedName");
                    _writer.WriteAttributeString("name", var);
                    _writer.WriteEndElement();
                }
                _writer.WriteEndElement();
            }

            if (sbd.XCTDataList.Count > 0)
            {
                _writer.WriteStartElement("sheetDataSet");
                int counter = 0;
                foreach (var var in sbd.XCTDataList)
                {
                    _writer.WriteStartElement("sheetData");
                    _writer.WriteAttributeString("sheetId", counter.ToString());
                    counter++;
                    foreach (var crn in var.CRNDataList)
                    {
                        _writer.WriteStartElement("row");
                        _writer.WriteAttributeString("r", (crn.rw + 1).ToString());
                        for (byte i = crn.colFirst; i <= crn.colLast; i++)
                        {
                            _writer.WriteStartElement("cell");
                            _writer.WriteAttributeString("r", ExcelHelperClass.intToABCString((int)i, (crn.rw + 1).ToString()));
                            if (crn.oper[i - crn.colFirst] is bool)
                            {
                                _writer.WriteAttributeString("t", "b");
                                if ((bool)crn.oper[i - crn.colFirst])
                                {
                                    _writer.WriteElementString("v", "1");
                                }
                                else
                                {
                                    _writer.WriteElementString("v", "0");
                                }
                            }
                            if (crn.oper[i - crn.colFirst] is double)
                            {
                                // _writer.WriteAttributeString("t", "b");
                                _writer.WriteElementString("v", Convert.ToString(crn.oper[i - crn.colFirst], CultureInfo.GetCultureInfo("en-US")));
                            }
                            if (crn.oper[i - crn.colFirst] is string)
                            {
                                _writer.WriteAttributeString("t", "str");
                                _writer.WriteElementString("v", crn.oper[i - crn.colFirst].ToString());
                            }


                            _writer.WriteEndElement();
                        }

                        _writer.WriteEndElement();
                    }

                    _writer.WriteEndElement();
                }
                _writer.WriteEndElement();
            }

            _writer.WriteEndElement(); 
            _writer.WriteEndElement();      // close worksheet
            _writer.WriteEndDocument();


            
            
            sbd.ExternalLinkId = this.xlsContext.SpreadDoc.WorkbookPart.GetExternalLinkPart().RelId;
            sbd.ExternalLinkRef = this.xlsContext.SpreadDoc.WorkbookPart.GetExternalLinkPart().RelIdToString;

            // close writer 
            _writer.Flush();
        }
    }
}
