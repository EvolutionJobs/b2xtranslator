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
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies information about a legend entry which has been changed from the 
    /// default legend entry settings and specifies the beginning of a collection of records as 
    /// defined by the Chart Sheet Substream ABNF.<br/>
    /// The collection of records specifies legend entry formatting. <br/>
    /// On a chart where the legend contains legend entries for the series and trendlines, 
    /// as defined in the legend overview, there MUST be zero or one instances of this record 
    /// in the sequence of records that conforms to the SERIESFORMAT rule.
    /// </summary>
    [BiffRecordAttribute(RecordType.LegendException)]
    public class LegendException : BiffRecord
    {
        public const RecordType ID = RecordType.LegendException;

        /// <summary>
        /// An unsigned integer that specifies the legend entry.<br/>
        /// This field has different interpretations depending on the content of the legend in the chart. 
        /// The legend overview specifies the types of content the legend can contain.<br/>
        /// <list>
        /// <item>In a chart where the legend contains legend entries for the series and trendlines, this field MUST be 0xFFFF. This record specifies the legend entry of the series or trendline that contains this record.</item>
        /// <item>In a chart where the legend contains legend entries for each data point in the chart, this field specifies the zero-based index of a legend entry in the legend, where 0x0000 is the legend entry for the first data point in the series.</item>
        /// <item>In a chart with a surface chart group, this field specifies the zero-based index of a legend entry in the legend, where 0x0000 is the legend entry for the lowest band of the surface chart group.</item>
        /// </list>
        /// </summary>
        public ushort iss;

        /// <summary>
        /// A bit that specifies whether the legend entry specified by iss has been deleted.
        /// </summary>
        public bool fDeleted;

        /// <summary>
        /// A bit that specifies whether the legend entry specified by iss has been formatted. <br/>
        /// If this field is 1, there MUST be a sequence of records that conforms to the 
        /// ATTACHEDLABEL rule in the Chart Sheet Substream ABNF following this record.
        /// </summary>
        public bool fLabel;

        public LegendException(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.iss = reader.ReadUInt16();
            var flags = reader.ReadUInt16();
            this.fDeleted = Utils.BitmaskToBool(flags, 0x1);
            this.fLabel = Utils.BitmaskToBool(flags, 0x2);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
