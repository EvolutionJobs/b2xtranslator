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

using System;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.SpreadsheetML;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class WindowMapping : AbstractOpenXmlMapping,
          IMapping<WindowSequence>
    {
        ExcelContext _xlsContext;
        ChartsheetPart _chartsheetPart;
        int _window1Id = 0;
        bool _isChartsheetSubstream;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="xlsContext">The excel context object</param>
        public WindowMapping(ExcelContext xlsContext, ChartsheetPart chartsheetPart, int window1Id, bool isChartsheetSubstream)
            : base(chartsheetPart.XmlWriter)
        {
            this._xlsContext = xlsContext;
            this._chartsheetPart = chartsheetPart;
            this._window1Id = window1Id;
            this._isChartsheetSubstream = isChartsheetSubstream;
        }

        #region IMapping<WindowSequence> Members

        public void Apply(WindowSequence windowSequence)
        {
            this._writer.WriteStartElement(Sml.Sheet.ElSheetView, Sml.Ns);

            this._writer.WriteAttributeString(Sml.Sheet.AttrTabSelected, windowSequence.Window2.fSelected ? "1" : "0");

            // zoomScale
            if (windowSequence.Scl != null)
            {
                // zoomScale is a percentage in the range 10-400
                int zoomScale = Math.Min(400, Math.Max(10, windowSequence.Scl.nscl * 100 / windowSequence.Scl.dscl));
                this._writer.WriteAttributeString(Sml.Sheet.AttrZoomScale, zoomScale.ToString());
            }

            this._writer.WriteAttributeString(Sml.Sheet.AttrWorkbookViewId, this._window1Id.ToString());
            // TODO: complete mapping, certain fields must be ignored when in chart sheet substream

            this._writer.WriteEndElement();
        }

        #endregion
    }
}
