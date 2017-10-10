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
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.DrawingML;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class DataLabelMapping : AbstractChartMapping,
          IMapping<ChartFormatsSequence>
    {
        SeriesFormatSequence _seriesFormatSequence;

        public DataLabelMapping(ExcelContext workbookContext, ChartContext chartContext, SeriesFormatSequence parentSeriesFormatSequence)
            : base(workbookContext, chartContext)
        {
            this._seriesFormatSequence = parentSeriesFormatSequence;
        }

        public void Apply(ChartFormatsSequence chartFormatSequence)
        {
            if (chartFormatSequence.AttachedLabelSequences.Count != 0)
            {
                var attachedLabelSequence = chartFormatSequence.AttachedLabelSequences[0];

                _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElDLbls, Dml.Chart.Ns);
                {
                    //<xsd:element name="dLbl" type="CT_DLbl" minOccurs="0" maxOccurs="unbounded"/>
                    // TODO

                    //<xsd:element name="dLblPos" type="CT_DLblPos" minOccurs="0" maxOccurs="1">
                    // TODO

                    //<xsd:element name="leaderLines" type="CT_ChartLines" minOccurs="0" maxOccurs="1">
                    // TODO

                    //<xsd:element name="numFmt" type="CT_NumFmt" minOccurs="0" maxOccurs="1">
                    // TODO

                    //<xsd:element name="showLeaderLines" type="CT_Boolean" minOccurs="0" maxOccurs="1">
                    // TODO

                    if (attachedLabelSequence != null)
                    {
                        if (attachedLabelSequence.Text != null)
                        {
                            //<xsd:element name="showLegendKey" type="CT_Boolean" minOccurs="0" maxOccurs="1">
                            if (attachedLabelSequence.Text.fShowKey)
                            {
                                writeValueElement("showLegendKey", "1");
                            }
                        }

                        if (attachedLabelSequence.DataLabExtContents != null)
                        {
                            //<xsd:element name="separator" type="xsd:string" minOccurs="0" maxOccurs="1">
                            _writer.WriteElementString(Dml.Chart.Prefix, "separator", Dml.Chart.Ns, attachedLabelSequence.DataLabExtContents.rgchSep.Value);

                            //<xsd:element name="showBubbleSize" type="CT_Boolean" minOccurs="0" maxOccurs="1">
                            if (attachedLabelSequence.DataLabExtContents.fBubSizes)
                            {
                                writeValueElement("showBubbleSize", "1");
                            }

                            //<xsd:element name="showCatName" type="CT_Boolean" minOccurs="0" maxOccurs="1">
                            if (attachedLabelSequence.DataLabExtContents.fCatName)
                            {
                                writeValueElement("showCatName", "1");
                            }

                            //<xsd:element name="showPercent" type="CT_Boolean" minOccurs="0" maxOccurs="1">
                            if (attachedLabelSequence.DataLabExtContents.fPercent)
                            {
                                writeValueElement("showPercent", "1");
                            }

                            //<xsd:element name="showSerName" type="CT_Boolean" minOccurs="0" maxOccurs="1">
                            if (attachedLabelSequence.DataLabExtContents.fSerName)
                            {
                                writeValueElement("showSerName", "1");
                            }

                            //<xsd:element name="showVal" type="CT_Boolean" minOccurs="0" maxOccurs="1">
                            if (attachedLabelSequence.DataLabExtContents.fValue)
                            {
                                writeValueElement("showVal", "1");
                            }
                        }

                    }
                    //<xsd:element name="spPr" type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1">
                    if (attachedLabelSequence.FrameSequence != null)
                    {
                        attachedLabelSequence.FrameSequence.Convert(new ShapePropertiesMapping(this.WorkbookContext, this.ChartContext));
                    }

                    //<xsd:element name="txPr" type="a:CT_TextBody" minOccurs="0" maxOccurs="1">
                    // TODO

                    //<xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
                    // TODO

                }
                _writer.WriteEndElement();
            }

        }
    }
}
