using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.DrawingML;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System.Globalization;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class CatMapping : AbstractChartMapping,
          IMapping<SeriesFormatSequence>
    {
        string _parentElement;

        public CatMapping(ExcelContext workbookContext, ChartContext chartContext, string parentElement)
            : base(workbookContext, chartContext)
        {
            this._parentElement = parentElement;
        }

        public void Apply(SeriesFormatSequence seriesFormatSequence)
        {
            // find BRAI record for categories
            foreach (AiSequence aiSequence in seriesFormatSequence.AiSequences)
            {
                if (aiSequence.BRAI.braiId == BRAI.BraiId.SeriesCategory)
                {
                    BRAI brai = aiSequence.BRAI;

                    // don't create a c:cat node for automatically generated category axis data!
                    if (brai.rt != BRAI.DataSource.Automatic)
                    {
                        // c:cat (or c:xVal for scatter and bubble charts)
                        _writer.WriteStartElement(Dml.Chart.Prefix, this._parentElement, Dml.Chart.Ns);
                        {
                            switch (brai.rt)
                            {
                                case BRAI.DataSource.Literal:
                                    // c:strLit
                                    _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElStrLit, Dml.Chart.Ns);
                                    {
                                        convertStringPoints(seriesFormatSequence);
                                    }
                                    _writer.WriteEndElement();
                                    break;
                                case BRAI.DataSource.Reference:
                                    // c:strRef
                                    _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElStrRef, Dml.Chart.Ns);
                                    {
                                        // c:f
                                        string formula = FormulaInfixMapping.mapFormula(brai.formula.formula, this.WorkbookContext);
                                        _writer.WriteElementString(Dml.Chart.Prefix, Dml.Chart.ElF, Dml.Chart.Ns, formula);

                                        // c:strCache
                                        _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElStrCache, Dml.Chart.Ns);
                                        {
                                            convertStringPoints(seriesFormatSequence);
                                        }
                                        _writer.WriteEndElement();
                                    }
                                    _writer.WriteEndElement();
                                    break;
                            }
                        }
                        _writer.WriteEndElement(); // c:cat
                    }
                    break;
                }
            }

        }

        private void convertStringPoints(SeriesFormatSequence seriesFormatSequence)
        {
            // find series data
            SeriesDataSequence seriesDataSequence = this.ChartContext.ChartSheetContentSequence.SeriesDataSequence;
            foreach (SeriesGroup seriesGroup in seriesDataSequence.SeriesGroups)
            {
                if (seriesGroup.SIIndex.numIndex == SIIndex.SeriesDataType.CategoryLabels)
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
                        AbstractCellContent cellContent = dataMatrix[seriesFormatSequence.order, i];
                        if (cellContent != null)
                        {
                            if (cellContent is Label)
                            {
                                Label lblInCell = (Label)cellContent;

                                // c:pt
                                _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElPt, Dml.Chart.Ns);
                                _writer.WriteAttributeString(Dml.Chart.AttrIdx, idx.ToString());

                                // c:v
                                _writer.WriteElementString(Dml.Chart.Prefix, Dml.Chart.ElV, Dml.Chart.Ns, lblInCell.st.Value);

                                _writer.WriteEndElement(); // c:pt
                            }
                        }
                        idx++;
                    }

                    break;
                }
            }
        }
    }
}
