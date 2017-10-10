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
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.SpreadsheetML;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class ChartsheetMapping : AbstractOpenXmlMapping,
          IMapping<ChartSheetSequence>
    {
        ExcelContext _xlsContext;
        ChartsheetPart _chartsheetPart;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="xlsContext">The excel context object</param>
        public ChartsheetMapping(ExcelContext xlsContext, ChartsheetPart targetPart)
            : base(targetPart.XmlWriter)
        {
            this._xlsContext = xlsContext;
            this._chartsheetPart = targetPart;
        }

        /// <summary>
        /// The overload apply method 
        /// Creates the Worksheet xml document 
        /// </summary>
        /// <param name="bsd">WorkSheetData</param>
        public void Apply(ChartSheetSequence chartSheetSequence)
        {
            this._writer.WriteStartDocument();
            // chartsheet
            this._writer.WriteStartElement(Sml.Sheet.ElChartsheet, Sml.Ns);
            this._writer.WriteAttributeString("xmlns", Sml.Ns);
            this._writer.WriteAttributeString("xmlns", "r", "", OpenXmlNamespaces.Relationships);
            
            var chartSheetContentSequence = chartSheetSequence.ChartSheetContentSequence;

            // sheetPr
            this._writer.WriteStartElement(Sml.Sheet.ElSheetPr, Sml.Ns);
            if (chartSheetContentSequence.CodeName != null)
            {
                // code name
                this._writer.WriteAttributeString(Sml.Sheet.AttrCodeName, chartSheetContentSequence.CodeName.codeName.Value);
            }
            // TODO: map SheetExtOptional to published and tab color

            this._writer.WriteEndElement();

            
            // sheetViews
            if (chartSheetContentSequence.WindowSequences.Count > 0)
            {
                this._writer.WriteStartElement(Sml.Sheet.ElSheetViews, Sml.Ns);

                // Note: There is a Window2 record for each Window1 record in the beginning of the workbook.
                // The index in the list corresponds to the 0-based workbookViewId attribute.
                //
                for (int window1Id = 0; window1Id < chartSheetContentSequence.WindowSequences.Count; window1Id++)
                {
                    chartSheetContentSequence.WindowSequences[window1Id].Convert(new WindowMapping(this._xlsContext, this._chartsheetPart, window1Id, true));
                }
                this._writer.WriteEndElement();
            }

            // page setup
            chartSheetContentSequence.PageSetupSequence.Convert(new PageSetupMapping(this._xlsContext, this._chartsheetPart));

            // header and footer
            // TODO: header and footer

            // drawing
            this._writer.WriteStartElement(Sml.Sheet.ElDrawing, Sml.Ns);
            this._writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, this._chartsheetPart.DrawingsPart.RelIdToString);
            chartSheetContentSequence.Convert(new DrawingMapping(this._xlsContext, this._chartsheetPart.DrawingsPart, true));

            this._writer.WriteEndElement();
            this._writer.WriteEndDocument();

            this._writer.Flush();
        }
    }
}
