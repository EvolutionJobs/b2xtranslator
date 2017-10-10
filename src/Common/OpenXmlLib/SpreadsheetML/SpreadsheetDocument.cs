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

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib.SpreadsheetML
{
    /// <summary>
    /// Includes some information about the spreadsheetdocument 
    /// </summary>
    public class SpreadsheetDocument : OpenXmlPackage
    {
        protected WorkbookPart workBookPart;
        protected OpenXmlPackage.DocumentType _documentType;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="fileName">Filename of the file which should be written</param>
        protected SpreadsheetDocument(string fileName, OpenXmlPackage.DocumentType type)
            : base(fileName)
        {
            switch (type)
            {
                case OpenXmlPackage.DocumentType.Document:
                    this.workBookPart = new WorkbookPart(this, SpreadsheetMLContentTypes.Workbook);
                    break;
                case OpenXmlPackage.DocumentType.MacroEnabledDocument:
                    this.workBookPart = new WorkbookPart(this, SpreadsheetMLContentTypes.WorkbookMacro);
                    break;
                //case OpenXmlPackage.DocumentType.Template:
                //    workBookPart = new WorkbookPart(this, WordprocessingMLContentTypes.MainDocumentTemplate);
                //    break;
                //case OpenXmlPackage.DocumentType.MacroEnabledTemplate:
                //    workBookPart = new WorkbookPart(this, WordprocessingMLContentTypes.MainDocumentMacroTemplate);
                //    break;
            }
            _documentType = type;
            this.AddPart(this.workBookPart);
        }

        /// <summary>
        /// creates a new excel document with the choosen filename 
        /// </summary>
        /// <param name="fileName">The name of the file which should be written</param>
        /// <returns>The object itself</returns>
        public static SpreadsheetDocument Create(string fileName, OpenXmlPackage.DocumentType type)
        {
            var spreadsheet = new SpreadsheetDocument(fileName, type);
            return spreadsheet;
        }

        public OpenXmlPackage.DocumentType DocumentType
        {
            get { return _documentType; }
            set { _documentType = value; }
        }

        /// <summary>
        /// returns the workbookPart from the new excel document 
        /// </summary>
        public WorkbookPart WorkbookPart
        {
            get { return this.workBookPart; }
        }
    }
}
