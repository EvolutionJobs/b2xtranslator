/*
 * Copyright (c) 2009, DIaLOGIKa
 *
 * All rights reserved.
 * 
 * Redistribution and use in source and binary forms, with or without 
 * modification, are permitted provided that the following conditions are met:
 * 
 *     * Redistributions of source code must retain the above copyright 
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright 
 *       notice, this list of conditions and the following disclaimer in the 
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the names of copyright holders, nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software 
 *       without specific prior written permission.
 * 
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" 
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED 
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. 
 * IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, 
 * INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, 
 * BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, 
 * DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF 
 * LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE 
 * OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF 
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE. 
 */

using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.DrawingML;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System;
using System.Globalization;
using System.Xml;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public abstract class AbstractChartMapping : AbstractOpenXmlMapping
    {
        private ExcelContext _workbookContext;
        private ChartContext _chartContext;
        
        public AbstractChartMapping(ExcelContext workbookContext, ChartContext chartContext)
            : base(chartContext.ChartPart.XmlWriter)
        {
            this._workbookContext = workbookContext;
            this._chartContext = chartContext;
        }

        public ExcelContext WorkbookContext
        {
            get { return this._workbookContext; }
        }

        public ChartContext ChartContext
        {
            get { return this._chartContext; }
        }

        public ChartPart ChartPart
        {
            get { return this.ChartContext.ChartPart; }
        }

        public ChartSheetContentSequence ChartSheetContentSequence
        {
            get { return this.ChartContext.ChartSheetContentSequence; }
        }

        public ChartContext.ChartLocation Location
        {
            get { return this.ChartContext.Location; }
        }

        public ChartFormatsSequence ChartFormatsSequence
        {
            get { return this.ChartSheetContentSequence.ChartFormatsSequence; }
        }

        protected void writeValueElement(string localName, string value)
        {
            _writer.WriteStartElement(Dml.Chart.Prefix, localName, Dml.Chart.Ns);
            _writer.WriteAttributeString(Dml.BaseTypes.AttrVal, value);
            _writer.WriteEndElement();
        }

        protected void writeValueElement(string prefix, string localName, string ns, string value)
        {
            _writer.WriteStartElement(prefix, localName, ns);
            _writer.WriteAttributeString(Dml.BaseTypes.AttrVal, value);
            _writer.WriteEndElement();
        }
    }
}
