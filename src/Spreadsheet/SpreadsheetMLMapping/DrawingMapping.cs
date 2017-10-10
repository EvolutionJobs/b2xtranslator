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
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.DrawingML;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class DrawingMapping : AbstractOpenXmlMapping,
        IMapping<ChartSheetContentSequence>,
        IMapping<ObjectsSequence>
    {
        ExcelContext _xlsContext;
        DrawingsPart _drawingsPart;

        bool _isChartsheet;

        public DrawingMapping(ExcelContext xlsContext, DrawingsPart targetPart, bool isChartsheet)
            : base(targetPart.XmlWriter)
        {
            this._xlsContext = xlsContext;
            this._drawingsPart = targetPart;

            this._isChartsheet = isChartsheet;
        }

        #region IMapping<ChartSheetContentSequence> Members

        public void Apply(ChartSheetContentSequence chartSheetContentSequence)
        {
            _writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElWsDr, Dml.SpreadsheetDrawing.Ns);
            _writer.WriteAttributeString("xmlns", Dml.SpreadsheetDrawing.Prefix, "", Dml.SpreadsheetDrawing.Ns);
            _writer.WriteAttributeString("xmlns", Dml.Prefix, "", Dml.Ns);

            if (this._isChartsheet)
            {
                _writer.WriteStartElement(Dml.SpreadsheetDrawing.ElAbsoluteAnchor, Dml.SpreadsheetDrawing.Ns);
                {
                    var chart = chartSheetContentSequence.ChartFormatsSequence.Chart;

                    // NOTE: Excel seems to somehow round the pos and ext values. The exact calculation is not documented.
                    //   Besides, Excel might write negative values which are corrected to 0 by Excel on load time.
                    //
                    // xdr:pos
                    _writer.WriteStartElement(Dml.SpreadsheetDrawing.ElPos, Dml.SpreadsheetDrawing.Ns);
                    _writer.WriteAttributeString(Dml.BaseTypes.AttrX, Math.Max(0, new PtValue(chart.x.Value).ToEmu()).ToString());
                    _writer.WriteAttributeString(Dml.BaseTypes.AttrY, Math.Max(0, new PtValue(chart.y.Value).ToEmu()).ToString());
                    _writer.WriteEndElement();

                    // xdr:ext
                    _writer.WriteStartElement(Dml.SpreadsheetDrawing.ElExt, Dml.SpreadsheetDrawing.Ns);
                    _writer.WriteAttributeString(Dml.BaseTypes.AttrCx, Math.Max(0, new PtValue(chart.dx.Value).ToEmu()).ToString());
                    _writer.WriteAttributeString(Dml.BaseTypes.AttrCy, Math.Max(0, new PtValue(chart.dy.Value).ToEmu()).ToString());
                    _writer.WriteEndElement();

                    // insert EG_ObjectChoices type
                    insertObjectChoices(chartSheetContentSequence);
                }
                _writer.WriteEndElement(); // absoluteAnchor
            }
            else
            {
                // embedded drawing
            }

            _writer.WriteEndElement();
            _writer.WriteEndDocument();

            _writer.Flush();
        }

        #endregion

        #region IMapping<ObjectsSequence> Members

        /// <summary>
        /// Mapping definition for embedded objects
        /// </summary>
        /// <param name="objectsSequence"></param>
        public void Apply(ObjectsSequence objectsSequence)
        {
            _writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElWsDr, Dml.SpreadsheetDrawing.Ns);
            _writer.WriteAttributeString("xmlns", Dml.SpreadsheetDrawing.Prefix, "", Dml.SpreadsheetDrawing.Ns);
            _writer.WriteAttributeString("xmlns", Dml.Prefix, "", Dml.Ns);

            foreach (var drawingsGroup in objectsSequence.DrawingsGroup)
            {
                // TODO: currently only embedded charts are mapped. Shapes and images are not yet implemented.
                //    The check on the type of object would have to be removed here once shapes and images are supported.
                //
                var objGroup = drawingsGroup.Objects.Find(p => p.ChartSheetSequence != null);
                if (objGroup != null)
                {
                    var msoDrawing = drawingsGroup.MsoDrawing;

                    // find OfficeArtClientAnchorSheet with drawing
                    var container = msoDrawing.rgChildRec as RegularContainer;

                    if (container != null)
                    {
                        var clientAnchor = container.FirstDescendantWithType<ClientAnchor>();
                        if (clientAnchor != null)
                        {
                            var oartClientAnchor = new OfficeArtClientAnchorSheet(clientAnchor.RawData);

                            // xdr:twoCellAnchor
                            _writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElTwoCellAnchor, Dml.SpreadsheetDrawing.Ns);
                            string editAs = "absolute";
                            if (oartClientAnchor.fSize && oartClientAnchor.fMove)
                            {
                                // Move and resize with anchor cells
                                editAs = "twoCell";
                            }
                            else if (!oartClientAnchor.fSize && oartClientAnchor.fMove)
                            {
                                // Move with cells but do not resize
                                editAs = "oneCell";
                            }
                            _writer.WriteAttributeString(Dml.SpreadsheetDrawing.AttrEditAs, editAs);
                            {
                                // xdr:from
                                _writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElFrom, Dml.SpreadsheetDrawing.Ns);
                                _writer.WriteElementString(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElCol, Dml.SpreadsheetDrawing.Ns, oartClientAnchor.colL.ToString());
                                _writer.WriteElementString(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElColOff, Dml.SpreadsheetDrawing.Ns, oartClientAnchor.dxL.ToString());
                                _writer.WriteElementString(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElRow, Dml.SpreadsheetDrawing.Ns, oartClientAnchor.rwT.ToString());
                                _writer.WriteElementString(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElRowOff, Dml.SpreadsheetDrawing.Ns, oartClientAnchor.dyT.ToString());
                                _writer.WriteEndElement(); // xdr:from

                                // xdr:to
                                _writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElTo, Dml.SpreadsheetDrawing.Ns);
                                _writer.WriteElementString(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElCol, Dml.SpreadsheetDrawing.Ns, oartClientAnchor.colR.ToString());
                                _writer.WriteElementString(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElColOff, Dml.SpreadsheetDrawing.Ns, oartClientAnchor.dxR.ToString());
                                _writer.WriteElementString(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElRow, Dml.SpreadsheetDrawing.Ns, oartClientAnchor.rwB.ToString());
                                _writer.WriteElementString(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElRowOff, Dml.SpreadsheetDrawing.Ns, oartClientAnchor.dyB.ToString());
                                _writer.WriteEndElement(); // xdr:to

                                var objectGroup = drawingsGroup.Objects.Find(p => p.ChartSheetSequence != null);
                                if (objectGroup != null)
                                {
                                    var chartSheetContentSequence = objectGroup.ChartSheetSequence.ChartSheetContentSequence;
                                    insertObjectChoices(chartSheetContentSequence);
                                }
                            }
                            _writer.WriteEndElement(); // xdr:twoCellAnchor
                        }
                    }
                }
            }

            _writer.WriteEndElement();
            _writer.WriteEndDocument();

            _writer.Flush();
        }

        #endregion

        private void insertObjectChoices(ChartSheetContentSequence chartSheetContentSequence)
        {
            _writer.WriteStartElement(Dml.SpreadsheetDrawing.ElGraphicFrame, Dml.SpreadsheetDrawing.Ns);
            {
                // TODO: add graphic properties
                _writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElNvGraphicFramePr, Dml.SpreadsheetDrawing.Ns);
                {
                    _writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElCNvPr, Dml.SpreadsheetDrawing.Ns);
                    _writer.WriteAttributeString(Dml.DocumentProperties.AttrId, this._drawingsPart.RelId.ToString());
                    _writer.WriteAttributeString(Dml.DocumentProperties.AttrName, "Shape");
                    _writer.WriteEndElement(); // xdr:cNvPr

                    _writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElCNvGraphicFramePr, Dml.SpreadsheetDrawing.Ns);
                    _writer.WriteStartElement(Dml.Prefix, Dml.DocumentProperties.ElGraphicFrameLocks, Dml.Ns);
                    _writer.WriteAttributeString(Dml.DocumentProperties.AttrNoGrp, "1");
                    _writer.WriteEndElement(); // a:graphicFrameLocks
                    _writer.WriteEndElement(); // xdr:cNvGraphicFramePr
                }
                _writer.WriteEndElement(); // xdr:nvGraphicFramePr

                // xdr:xfrm
                _writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElXfrm, Dml.SpreadsheetDrawing.Ns);
                {
                    _writer.WriteStartElement(Dml.Prefix, Dml.BaseTypes.ElOff, Dml.Ns);
                    _writer.WriteAttributeString(Dml.BaseTypes.AttrX, "0");
                    _writer.WriteAttributeString(Dml.BaseTypes.AttrY, "0");
                    _writer.WriteEndElement(); // a:off

                    _writer.WriteStartElement(Dml.Prefix, Dml.BaseTypes.ElExt, Dml.Ns);
                    _writer.WriteAttributeString(Dml.BaseTypes.AttrCx, "0");
                    _writer.WriteAttributeString(Dml.BaseTypes.AttrCy, "0");
                    _writer.WriteEndElement(); // a:ext
                }
                _writer.WriteEndElement(); // xdr:xfrm


                _writer.WriteStartElement(Dml.GraphicalObject.ElGraphic, Dml.Ns);
                {
                    _writer.WriteStartElement(Dml.GraphicalObject.ElGraphicData, Dml.Ns);
                    _writer.WriteAttributeString(Dml.GraphicalObject.AttrUri, Dml.Chart.Ns);

                    // create and convert chart part
                    var chartPart = this._drawingsPart.AddChartPart();
                    var chartContext = new ChartContext(chartPart, chartSheetContentSequence, this._isChartsheet ? ChartContext.ChartLocation.Chartsheet : ChartContext.ChartLocation.Embedded);
                    chartSheetContentSequence.Convert(new ChartMapping(this._xlsContext, chartContext));

                    _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElChart, Dml.Chart.Ns);
                    _writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, chartPart.RelIdToString);

                    _writer.WriteEndElement(); // c:chart

                    _writer.WriteEndElement(); // a:graphicData
                }
                _writer.WriteEndElement(); // a:graphic
            }
            _writer.WriteEndElement(); // a:graphicFrame

            _writer.WriteElementString(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElClientData, Dml.SpreadsheetDrawing.Ns, string.Empty);
        }
    }
}
