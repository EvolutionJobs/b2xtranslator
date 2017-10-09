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

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class ScatterChartMapping : AbstractChartGroupMapping
    {
        public ScatterChartMapping(ExcelContext workbookContext, ChartContext chartContext, bool is3DChart)
            : base(workbookContext, chartContext, is3DChart)
        {
        }

        #region IMapping<CrtSequence> Members

        public override void Apply(CrtSequence crtSequence)
        {
            if (!(crtSequence.ChartType is Scatter))
            {
                throw new Exception("Invalid chart type");
            }

            Scatter scatter = crtSequence.ChartType as Scatter;

            // c:scatterChart 
            _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElScatterChart, Dml.Chart.Ns);
            {
                // c:scatterStyle
                writeValueElement(Dml.Chart.Prefix, Dml.Chart.ElScatterStyle, Dml.Chart.Ns, mapScatterStyle(crtSequence.SsSequence));

                // c:varyColors
                //writeValueElement(Dml.Chart.ElVaryColors, crtSequence.ChartFormat.fVaried ? "1" : "0");

                foreach (SeriesFormatSequence seriesFormatSequence in this.ChartFormatsSequence.SeriesFormatSequences)
                {
                    if (seriesFormatSequence.SerToCrt != null && seriesFormatSequence.SerToCrt.id == crtSequence.ChartFormat.idx)
                    {
                        // c:ser (CT_ScatterSer)
                        // c:ser
                        _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElSer, Dml.Chart.Ns);

                        // EG_SerShared
                        seriesFormatSequence.Convert(new SeriesMapping(this.WorkbookContext, this.ChartContext));

                        // c:marker

                        // c:dPt
                        
                        // c:dLbls (CT_DLbls)
                        this.ChartFormatsSequence.Convert(new DataLabelMapping(this.WorkbookContext, this.ChartContext, seriesFormatSequence));

                        // c:trendline

                        // c:errBars

                        // c:xVal
                        seriesFormatSequence.Convert(new CatMapping(this.WorkbookContext, this.ChartContext, Dml.Chart.ElXVal));

                        // c:yVal
                        seriesFormatSequence.Convert(new ValMapping(this.WorkbookContext, this.ChartContext, Dml.Chart.ElYVal));

                        // c:smooth
                        writeValueElement(Dml.Chart.Prefix, Dml.Chart.ElSmooth, Dml.Chart.Ns, isSmoothed(crtSequence.SsSequence) ? "1" : "0");

                        _writer.WriteEndElement(); // c:ser
                    }
                }

                // Data Labels

                // Axis Ids
                foreach (int axisId in crtSequence.ChartFormat.AxisIds)
                {
                    writeValueElement(Dml.Chart.ElAxId, axisId.ToString());
                }
            }
            _writer.WriteEndElement();
        }
        #endregion

        private string mapScatterStyle(SsSequence ssSequence)
        {
            // CT_ScatterStyle
            // The following scatter styles exist: line, lineMarker, marker, none, smooth, smoothMarker
            // 
            bool smoothed = isSmoothed(ssSequence);
            bool hasMarker = (ssSequence == null) || (ssSequence.MarkerFormat != null && ssSequence.MarkerFormat.imk != MarkerFormat.MarkerType.NoMarker);
            bool hasLine = (ssSequence == null) || (ssSequence.LineFormat != null && ssSequence.LineFormat.lns != LineFormat.LineStyle.None);

            string scatterStyle = "none";
            if (smoothed && hasMarker)
            {
                scatterStyle = "smoothMarker";
            }
            else if (smoothed)
            {
                scatterStyle = "smooth";
            }
            else if (hasMarker && hasLine)
            {
                scatterStyle = "lineMarker";
            }
            else if (hasLine)
            {
                scatterStyle = "line";
            }
            else if (hasMarker)
            {
                scatterStyle = "marker";
            }
            return scatterStyle;
        }

        private bool isSmoothed(SsSequence ssSequence)
        {
            return (ssSequence != null && ssSequence.SerFmt != null && ssSequence.SerFmt.fSmoothedLine);
        }
    }
}
