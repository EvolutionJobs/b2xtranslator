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
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class ChartMapping : AbstractChartMapping,
          IMapping<ChartSheetContentSequence>
    {
        ChartContext _chartContext;

        public ChartMapping(ExcelContext workbookContext, ChartContext chartContext)
            : base(workbookContext, chartContext)
        {
            this._chartContext = chartContext;
        }

        #region IMapping<ChartSheetContentSequence> Members

        // TODO: maybe we only need chartSheetContentSequence.ChartFormatsSequence here
        public void Apply(ChartSheetContentSequence chartSheetContentSequence)
        {
            var chartFormatsSequence = chartSheetContentSequence.ChartFormatsSequence;

            // c:chartspace
            _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElChartSpace, Dml.Chart.Ns);
            _writer.WriteAttributeString("xmlns", Dml.Chart.Prefix, "", Dml.Chart.Ns);
            _writer.WriteAttributeString("xmlns", Dml.Prefix, "", Dml.Ns);
            {
                // c:chart
                _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElChart, Dml.Chart.Ns);
                {
                    // c:title
                    foreach (var attachedLabelSequence in chartFormatsSequence.AttachedLabelSequences)
                    {
                        if (attachedLabelSequence.ObjectLink != null && attachedLabelSequence.ObjectLink.wLinkObj == ObjectLink.ObjectType.Chart)
                        {
                            attachedLabelSequence.Convert(new TitleMapping(this.WorkbookContext, this.ChartContext));
                            break;
                        }
                    }

                    // c:plotArea
                    chartFormatsSequence.Convert(new PlotAreaMapping(this.WorkbookContext, this.ChartContext));

                    // c:legend
                    var firstLegend = chartFormatsSequence.AxisParentSequences[0].CrtSequences[0].LdSequence;
                    if (firstLegend != null)
                    {
                        firstLegend.Convert(new LegendMapping(this.WorkbookContext, this.ChartContext));
                    }

                    // c:plotVisOnly
                    writeValueElement(Dml.Chart.ElPlotVisOnly, chartFormatsSequence.ShtProps.fPlotVisOnly ? "1" : "0");
                    
                    // c:dispBlanksAs
                    string dispBlanksAs = string.Empty;
                    switch (chartFormatsSequence.ShtProps.mdBlank)
                    {
                        case ShtProps.EmptyCellPlotMode.PlotNothing:
                            dispBlanksAs = "gap";
                            break;

                        case ShtProps.EmptyCellPlotMode.PlotAsZero:
                            dispBlanksAs = "zero";
                            break;

                        case ShtProps.EmptyCellPlotMode.PlotAsInterpolated:
                            dispBlanksAs = "span";
                            break;
                    }
                    if (!string.IsNullOrEmpty(dispBlanksAs))
                    {
                        writeValueElement(Dml.Chart.ElDispBlanksAs, dispBlanksAs);
                    }

                    // c:showDLblsOverMax

                }
                _writer.WriteEndElement(); // c:chart

                // c:spPr
                if (chartFormatsSequence.FrameSequence != null)
                {
                    chartFormatsSequence.FrameSequence.Convert(new ShapePropertiesMapping(this.WorkbookContext, this.ChartContext));
                }
            }
            _writer.WriteEndElement();
            _writer.WriteEndDocument();

            _writer.Flush();
        }

        #endregion
    }
}
