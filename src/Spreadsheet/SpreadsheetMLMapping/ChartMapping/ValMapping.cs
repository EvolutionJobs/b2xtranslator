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
using System;
using System.Globalization;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class ValMapping : AbstractChartMapping,
          IMapping<SeriesFormatSequence>
    {
        string _parentElement;

        public ValMapping(ExcelContext workbookContext, ChartContext chartContext, string parentElement)
            : base(workbookContext, chartContext)
        {
            this._parentElement = parentElement;
        }

        #region IMapping<SeriesFormatSequence> Members

        /// <summary>
        /// </summary>
        public void Apply(SeriesFormatSequence seriesFormatSequence)
        {
            // find BRAI record for values
            foreach (AiSequence aiSequence in seriesFormatSequence.AiSequences)
            {
                if (aiSequence.BRAI.braiId == BRAI.BraiId.SeriesValues)
                {
                    // c:val
                    _writer.WriteStartElement(Dml.Chart.Prefix, this._parentElement, Dml.Chart.Ns);
                    {
                        BRAI brai = aiSequence.BRAI;
                        switch (brai.rt)
                        {
                            case BRAI.DataSource.Literal:
                                // c:numLit
                                _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElNumLit, Dml.Chart.Ns);

                                convertNumData(seriesFormatSequence);

                                _writer.WriteEndElement(); // c:numLit
                                break;

                            case BRAI.DataSource.Reference:
                                // c:numRef
                                _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElNumRef, Dml.Chart.Ns);

                                // c:f
                                string formula = FormulaInfixMapping.mapFormula(brai.formula.formula, this.WorkbookContext);
                                _writer.WriteElementString(Dml.Chart.Prefix, Dml.Chart.ElF, Dml.Chart.Ns, formula);

                                // TODO: optional data cache
                                // c:numCache
                                _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElNumCache, Dml.Chart.Ns);
                                convertNumData(seriesFormatSequence);
                                _writer.WriteEndElement(); // c:numCache

                                _writer.WriteEndElement(); // c:numRef
                                break;
                        }
                    }
                    _writer.WriteEndElement(); // c:val
                    break;
                }
            }
        }
        #endregion

        private void convertNumData(SeriesFormatSequence seriesFormatSequence)
        {
            // find series data
            SeriesDataSequence seriesDataSequence = this.ChartContext.ChartSheetContentSequence.SeriesDataSequence;
            foreach (SeriesGroup seriesGroup in seriesDataSequence.SeriesGroups)
            {
                if (seriesGroup.SIIndex.numIndex == SIIndex.SeriesDataType.SeriesValues)
                {
                    AbstractCellContent[,] dataMatrix = seriesDataSequence.DataMatrix[(UInt16)seriesGroup.SIIndex.numIndex - 1];
                    // TODO: c:formatCode

                    UInt32 ptCount = 0;
                    for (UInt32 i = 0; i < dataMatrix.GetLength(1); i++)
                    {
                        if (dataMatrix[seriesFormatSequence.order, i] != null)
                        {
                            ptCount++;
                        }
                    }
                    // c:ptCount
                    writeValueElement(Dml.Chart.ElPtCount, ptCount.ToString());

                    UInt32 idx = 0;
                    for (UInt32 i = 0; i < dataMatrix.GetLength(1); i++)
                    {
                        Number cellContent = dataMatrix[seriesFormatSequence.order, i] as Number;
                        if (cellContent != null && cellContent.num != null)
                        {
                            // c:pt
                            _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElPt, Dml.Chart.Ns);
                            _writer.WriteAttributeString(Dml.Chart.AttrIdx, idx.ToString());

                            // c:v
                            double num = cellContent.num ?? 0.0;
                            _writer.WriteElementString(Dml.Chart.Prefix, Dml.Chart.ElV, Dml.Chart.Ns, num.ToString(CultureInfo.InvariantCulture));

                            _writer.WriteEndElement(); // c:pt
                        }
                        idx++;
                    }

                    break;
                }
            }
        }
    }
}
