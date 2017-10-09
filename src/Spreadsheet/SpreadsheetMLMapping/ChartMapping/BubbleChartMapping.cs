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
    public class BubbleChartMapping : AbstractChartGroupMapping
    {
        public BubbleChartMapping(ExcelContext workbookContext, ChartContext chartContext, bool is3DChart)
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

            // c:bubbleChart 
            _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElBubbleChart, Dml.Chart.Ns);
            {
                // c:varyColors: This setting needs to be ignored if the chart has 
                //writeValueElement(Dml.Chart.ElVaryColors, crtSequence.ChartFormat.fVaried ? "1" : "0");

                // Bubble Chart Series
                foreach (SeriesFormatSequence seriesFormatSequence in this.ChartFormatsSequence.SeriesFormatSequences)
                {
                    if (seriesFormatSequence.SerToCrt != null && seriesFormatSequence.SerToCrt.id == crtSequence.ChartFormat.idx)
                    {
                        // c:ser
                        _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElSer, Dml.Chart.Ns);

                        // EG_SerShared
                        seriesFormatSequence.Convert(new SeriesMapping(this.WorkbookContext, this.ChartContext));

                        // c:dPt

                        // c:dLbls (CT_DLbls)
                        this.ChartFormatsSequence.Convert(new DataLabelMapping(this.WorkbookContext, this.ChartContext, seriesFormatSequence));

                        // c:trendline

                        // c:errBars

                        // c:xVal
                        seriesFormatSequence.Convert(new CatMapping(this.WorkbookContext, this.ChartContext, Dml.Chart.ElXVal));

                        // c:yVal
                        seriesFormatSequence.Convert(new ValMapping(this.WorkbookContext, this.ChartContext, Dml.Chart.ElYVal));

                        // c:bubbleSize

                        // c:bubble3D

                        _writer.WriteEndElement(); // c:ser
                    }
                }

                // c:dLbls

                // c:bubble3D

                // c:bubbleScale

                // c:showNegBubbles

                // c:sizeRepresents


                // Axis Ids
                foreach (int axisId in crtSequence.ChartFormat.AxisIds)
                {
                    writeValueElement(Dml.Chart.ElAxId, axisId.ToString());
                }
            }
            _writer.WriteEndElement();
        }
        #endregion
    }
}
