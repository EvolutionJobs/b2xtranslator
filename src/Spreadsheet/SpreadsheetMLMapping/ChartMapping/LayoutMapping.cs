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
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System.Globalization;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class LayoutMapping : AbstractChartMapping,
          IMapping<CrtLayout12>
    {
        public LayoutMapping(ExcelContext workbookContext, ChartContext chartContext)
            : base(workbookContext, chartContext)
        {
        }

        #region IMapping<CrtLayout12> Members

        public void Apply(CrtLayout12 crtLayout12)
        {
            // c:layout
            this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElLayout, Dml.Chart.Ns);
            {
                if (crtLayout12.wHeightMode != CrtLayout12.CrtLayout12Mode.L12MAUTO ||
                    crtLayout12.wWidthMode != CrtLayout12.CrtLayout12Mode.L12MAUTO ||
                    crtLayout12.wYMode != CrtLayout12.CrtLayout12Mode.L12MAUTO ||
                    crtLayout12.wXMode != CrtLayout12.CrtLayout12Mode.L12MAUTO)
                {
                    // c:manualLayout
                    this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElManualLayout, Dml.Chart.Ns);
                    {
                        // c:layoutTarget
                        writeValueElement(Dml.Chart.ElLayoutTarget, crtLayout12.fLayoutTargetInner ? "inner" : "outer");

                        if (crtLayout12.wXMode != CrtLayout12.CrtLayout12Mode.L12MAUTO)
                        {
                            // c:xMode
                            writeValueElement(Dml.Chart.ElXMode, crtLayout12.wXMode == CrtLayout12.CrtLayout12Mode.L12MEDGE ? "edge" : "factor");
                        }
                        if (crtLayout12.wYMode != CrtLayout12.CrtLayout12Mode.L12MAUTO)
                        {
                            // c:yMode
                            writeValueElement(Dml.Chart.ElYMode, crtLayout12.wYMode == CrtLayout12.CrtLayout12Mode.L12MEDGE ? "edge" : "factor");
                        }
                        if (crtLayout12.wWidthMode != CrtLayout12.CrtLayout12Mode.L12MAUTO)
                        {
                            // c:wMode
                            writeValueElement(Dml.Chart.ElWMode, crtLayout12.wWidthMode == CrtLayout12.CrtLayout12Mode.L12MEDGE ? "edge" : "factor");
                        }
                        if (crtLayout12.wHeightMode != CrtLayout12.CrtLayout12Mode.L12MAUTO)
                        {
                            // c:hMode
                            writeValueElement(Dml.Chart.ElHMode, crtLayout12.wHeightMode == CrtLayout12.CrtLayout12Mode.L12MEDGE ? "edge" : "factor");
                        }

                        // c:x
                        writeValueElement(Dml.Chart.ElX, crtLayout12.x.ToString(CultureInfo.InvariantCulture));

                        // c:y
                        writeValueElement(Dml.Chart.ElY, crtLayout12.y.ToString(CultureInfo.InvariantCulture));

                        // c:w
                        writeValueElement(Dml.Chart.ElW, crtLayout12.dx.ToString(CultureInfo.InvariantCulture));

                        // c:h
                        writeValueElement(Dml.Chart.ElH, crtLayout12.dy.ToString(CultureInfo.InvariantCulture));
                    }
                    this._writer.WriteEndElement(); // c:manualLayout
                }
            }
            this._writer.WriteEndElement(); // c:layout
        }
        #endregion
    }
}
