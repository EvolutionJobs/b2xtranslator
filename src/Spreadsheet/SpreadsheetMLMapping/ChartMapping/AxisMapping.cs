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
    public class AxisMapping : AbstractChartMapping,
          IMapping<AxesSequence>
    {
        public AxisMapping(ExcelContext workbookContext, ChartContext chartContext)
            : base(workbookContext, chartContext)
        {
        }

        #region IMapping<AxesSequence> Members

        public virtual void Apply(AxesSequence axesSequence)
        {
            if (axesSequence != null)
            {
                if (axesSequence.IvAxisSequence != null)
                {
                    // c:catAx
                    _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElCatAx, Dml.Chart.Ns);
                    {
                        mapIvAxis(axesSequence.IvAxisSequence, axesSequence);
                    }
                    _writer.WriteEndElement(); // c:catAx
                }


                // c:valAx
                _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElValAx, Dml.Chart.Ns);
                {
                    // c:axId
                    writeValueElement(Dml.Chart.ElAxId, axesSequence.DvAxisSequence.Axis.AxisId.ToString());

                    // c:scaling
                    mapScaling(axesSequence.DvAxisSequence.ValueRange);

                    // c:delete

                    // c:axPos
                    // TODO: find mapping
                    writeValueElement(Dml.Chart.ElAxPos, "l");

                    // c:majorGridlines

                    // c:minorGridlines

                    // c:title
                    foreach (AttachedLabelSequence attachedLabelSequence in axesSequence.AttachedLabelSequences)
                    {
                        if (attachedLabelSequence.ObjectLink != null && attachedLabelSequence.ObjectLink.wLinkObj == ObjectLink.ObjectType.DVAxis)
                        {
                            attachedLabelSequence.Convert(new TitleMapping(this.WorkbookContext, this.ChartContext));
                            break;
                        }
                    }

                    // c:numFmt

                    // c:majorTickMark

                    // c:minorTickMark

                    // c:tickLblPos

                    // c:spPr

                    // c:txPr

                    // c:crossAx
                    if (axesSequence.IvAxisSequence != null)
                    {
                        writeValueElement(Dml.Chart.ElCrossAx, axesSequence.IvAxisSequence.Axis.AxisId.ToString());
                    }
                    else if (axesSequence.DvAxisSequence2 != null)
                    {
                        writeValueElement(Dml.Chart.ElCrossAx, axesSequence.DvAxisSequence2.Axis.AxisId.ToString());
                    }

                    // c:crosses or c:crossesAt

                    // c:crossBetween

                    // c:majorUnit
                    mapMajorUnit(axesSequence.DvAxisSequence.ValueRange);

                    // c:minorUnit
                    mapMinorUnit(axesSequence.DvAxisSequence.ValueRange);

                    // c:dispUnits

                }
                _writer.WriteEndElement(); // c:valAx

                if (axesSequence.DvAxisSequence2 != null)
                {
                    _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElValAx, Dml.Chart.Ns);
                    {
                        // c:axId
                        writeValueElement(Dml.Chart.ElAxId, axesSequence.DvAxisSequence2.Axis.AxisId.ToString());

                        // c:scaling
                        mapScaling(axesSequence.DvAxisSequence2.ValueRange);

                        // c:delete

                        // c:axPos
                        // TODO: find mapping
                        writeValueElement(Dml.Chart.ElAxPos, "b");

                        // c:majorGridlines

                        // c:minorGridlines

                        // c:title
                        foreach (AttachedLabelSequence attachedLabelSequence in axesSequence.AttachedLabelSequences)
                        {
                            if (attachedLabelSequence.ObjectLink != null && attachedLabelSequence.ObjectLink.wLinkObj == ObjectLink.ObjectType.DVAxis)
                            {
                                attachedLabelSequence.Convert(new TitleMapping(this.WorkbookContext, this.ChartContext));
                                break;
                            }
                        }

                        // c:numFmt

                        // c:majorTickMark

                        // c:minorTickMark

                        // c:tickLblPos

                        // c:spPr

                        // c:txPr

                        // c:crossAx
                        if (axesSequence.DvAxisSequence != null)
                        {
                            writeValueElement(Dml.Chart.ElCrossAx, axesSequence.DvAxisSequence.Axis.AxisId.ToString());
                        }

                        // c:crosses or c:crossesAt

                        // c:crossBetween

                        // c:majorUnit
                        mapMajorUnit(axesSequence.DvAxisSequence2.ValueRange);

                        // c:minorUnit
                        mapMinorUnit(axesSequence.DvAxisSequence2.ValueRange);

                        // c:dispUnits

                    }
                    _writer.WriteEndElement(); // c:valAx
                }

                if (axesSequence.SeriesAxisSequence != null)
                {
                    // c:serAx
                    _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElSerAx, Dml.Chart.Ns);
                    {
                        mapIvAxis(axesSequence.SeriesAxisSequence, axesSequence);
                    }
                    _writer.WriteEndElement(); // c:serAx
                }
            }
        }

        #endregion

        private void mapIvAxis(IvAxisSequence ivAxisSequence, AxesSequence axesSequence)
        {
            // EG_AxShared
            // c:axId
            writeValueElement(Dml.Chart.ElAxId, ivAxisSequence.Axis.AxisId.ToString());

            // c:scaling
            _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElScaling, Dml.Chart.Ns);
            {
                // c:logBase

                // c:orientation
                writeValueElement(Dml.Chart.ElOrientation, ivAxisSequence.CatSerRange.fReverse ? "maxMin" : "minMax");

                // c:max

                // c:min


            }
            _writer.WriteEndElement(); // c:scaling

            // c:delete

            // c:axPos
            // TODO: find mapping
            writeValueElement(Dml.Chart.ElAxPos, "b");

            // c:majorGridlines

            // c:minorGridlines

            // c:title
            foreach (AttachedLabelSequence attachedLabelSequence in axesSequence.AttachedLabelSequences)
            {
                if (attachedLabelSequence.ObjectLink != null && attachedLabelSequence.ObjectLink.wLinkObj == ObjectLink.ObjectType.IVAxis)
                {
                    attachedLabelSequence.Convert(new TitleMapping(this.WorkbookContext, this.ChartContext));
                    break;
                }
            }

            // c:numFmt

            // c:majorTickMark

            // c:minorTickMark

            // c:tickLblPos

            // c:spPr

            // c:txPr

            // c:crossAx
            writeValueElement(Dml.Chart.ElCrossAx, axesSequence.DvAxisSequence.Axis.AxisId.ToString());

            // c:crosses or c:crossesAt

        }


        private void mapScaling(ValueRange valueRange)
        {
            // c:scaling
            _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElScaling, Dml.Chart.Ns);
            {
                // c:logBase
                // TODO: support for custom logarithmic base (The default base of the logarithmic scale is 10, 
                //  unless a CrtMlFrt record follows this record, specifying the base in a XmlTkLogBaseFrt)

                // c:orientation
                writeValueElement(Dml.Chart.ElOrientation, (valueRange == null || valueRange.fReversed) ? "maxMin" : "minMax");

                // c:max
                if (valueRange != null && !valueRange.fAutoMax)
                {
                    writeValueElement(Dml.Chart.ElMax, valueRange.numMax.ToString(CultureInfo.InvariantCulture));
                }

                // c:min
                if (valueRange != null && !valueRange.fAutoMin)
                {
                    writeValueElement(Dml.Chart.ElMin, valueRange.numMin.ToString(CultureInfo.InvariantCulture));
                }
            }
            _writer.WriteEndElement(); // c:scaling

        }

        private void mapMajorUnit(ValueRange valueRange)
        {
            if (valueRange != null && !valueRange.fAutoMajor)
            {
                writeValueElement(Dml.Chart.ElMajorUnit, valueRange.numMajor.ToString(CultureInfo.InvariantCulture));
            }
        }

        private void mapMinorUnit(ValueRange valueRange)
        {
            if (valueRange != null && !valueRange.fAutoMinor)
            {
                writeValueElement(Dml.Chart.ElMinorUnit, valueRange.numMinor.ToString(CultureInfo.InvariantCulture));
            }
        }
    }
}
