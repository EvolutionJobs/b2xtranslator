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
    public class SeriesMapping : AbstractChartMapping,
          IMapping<SeriesFormatSequence>
    {
        public SeriesMapping(ExcelContext workbookContext, ChartContext chartContext)
            : base(workbookContext, chartContext)
        {
        }

        #region IMapping<SeriesFormatSequence> Members

        public void Apply(SeriesFormatSequence seriesFormatSequence)
        {
            // EG_SerShared

            // c:idx
            // TODO: check the meaning of this element
            writeValueElement(Dml.Chart.ElIdx, seriesFormatSequence.SerToCrt.id.ToString());
            
            // c:order
            writeValueElement(Dml.Chart.ElOrder, seriesFormatSequence.order.ToString());

            // c:tx
            // find BRAI record for series name
            foreach (var aiSequence in seriesFormatSequence.AiSequences)
            {
                if (aiSequence.BRAI.braiId == BRAI.BraiId.SeriesNameOrLegendText)
                {
                    var brai = aiSequence.BRAI;
                    
                    if (aiSequence.SeriesText != null)
                    {
                        switch (brai.rt)
                        {
                            case BRAI.DataSource.Literal:
                                // c:tx
                                this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElTx, Dml.Chart.Ns);
                                // c:v
                                this._writer.WriteElementString(Dml.Chart.Prefix, Dml.Chart.ElV, Dml.Chart.Ns, aiSequence.SeriesText.stText.Value);
                                this._writer.WriteEndElement(); // c:tx
                                break;

                            case BRAI.DataSource.Reference:
                                // c:tx
                                this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElTx, Dml.Chart.Ns);


                                // c:strRef
                                this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElStrRef, Dml.Chart.Ns);
                                {
                                    // c:f
                                    string formula = FormulaInfixMapping.mapFormula(brai.formula.formula, this.WorkbookContext);
                                    this._writer.WriteElementString(Dml.Chart.Prefix, Dml.Chart.ElF, Dml.Chart.Ns, formula);

                                    // c:strCache
                                    //convertStringCache(seriesFormatSequence);
                                }

                                this._writer.WriteEndElement(); // c:strRef
                                this._writer.WriteEndElement(); // c:tx
                                break;
                        }
                    }

                    break;
                }
            }

            if (seriesFormatSequence.SsSequence.Count > 0)
            {
                seriesFormatSequence.SsSequence[0].Convert(new ShapePropertiesMapping(this.WorkbookContext, this.ChartContext));
            }

        }
        #endregion
    }
}
