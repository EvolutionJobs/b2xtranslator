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
using DIaLOGIKa.b2xtranslator.OpenXmlLib.SpreadsheetML;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    /// <summary>
    /// Includes some attributes and methods required by the mapping classes 
    /// </summary>
    public class ExcelContext
    {
        private SpreadsheetDocument spreadDoc;
        private XmlWriterSettings writerSettings;
        private XlsDocument xlsDoc;
        private SheetData currentSheet; 

        /// <summary>
        /// The settings of the XmlWriter which writes to the part
        /// </summary>
        public XmlWriterSettings WriterSettings
        {
            get { return writerSettings; }
            set { writerSettings = value; }
        }

        /// <summary>
        /// The XlsDocument 
        /// </summary>
        public SpreadsheetDocument SpreadDoc
        {
            get { return spreadDoc; }
            set { this.spreadDoc = value; }
        }

        /// <summary>
        /// The XlsDocument 
        /// </summary>
        public XlsDocument XlsDoc
        {
            get { return xlsDoc; }
            set { this.xlsDoc = value; }
        }

        /// <summary>
        /// Current working sheet !! !
        /// </summary>
        public SheetData CurrentSheet
        {
            get { return this.currentSheet; }
            set { this.currentSheet = value; }
        }

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="xlsDoc">Xls document </param>
        /// <param name="writerSettings">the xml writer settings </param>
        public ExcelContext(XlsDocument xlsDoc, XmlWriterSettings writerSettings)
        {
            this.xlsDoc = xlsDoc;
            this.writerSettings = writerSettings; 
        }
    }


}
