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

using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies properties of the data for a series, a trendline, or error bars, and 
    /// specifies the beginning of a collection of records as defined by the Chart Sheet Substream ABNF.<br/> 
    /// The collection of records specifies a series, a trendline, or error bars.
    /// </summary>
    [BiffRecordAttribute(RecordType.Series)]
    public class Series : BiffRecord
    {
        public enum SeriesDataType
        {
            Text = 3,
            Numeric = 1
        }

        public const RecordType ID = RecordType.Series;

        /// <summary>
        /// specifies the type of data in categories (3), or horizontal values on 
        /// bubble and scatter chart groups, in the series.
        /// </summary>
        public SeriesDataType sdtX;

        /// <summary>
        /// An unsigned integer that specifies that the values, or vertical values on bubble and 
        /// scatter chart groups, in the series contain numeric information. 
        /// It MUST be Numeric, and MUST be ignored.
        /// </summary>
        public SeriesDataType sdtY;

        /// <summary>
        /// An unsigned integer that specifies the count of categories (3), 
        /// or horizontal values on bubble and scatter chart groups, in the series. <br/>
        /// The value MUST be less than or equal to 3999.
        /// </summary>
        public ushort cValx;

        /// <summary>
        /// An unsigned integer that specifies the count of values, or vertical 
        /// values on bubble and scatter chart groups, in the series. <br/>
        /// The value MUST be less than or equal to 3999.
        /// </summary>
        public ushort cValy;

        /// <summary>
        /// specifies that the bubble size values in the series contain numeric information. 
        /// The value MUST be Numeric, and MUST be ignored.
        /// </summary>
        public SeriesDataType sdtBSize;

        /// <summary>
        /// An unsigned integer that specifies the count of bubble size values in the series. <br/>
        /// The value MUST be less than or equal to 3999.
        /// </summary>
        public ushort cValBSize;

        public Series(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.sdtX = (SeriesDataType)reader.ReadUInt16();
            this.sdtY = (SeriesDataType)reader.ReadUInt16();
            this.cValx = reader.ReadUInt16();
            this.cValy = reader.ReadUInt16();
            this.sdtBSize = (SeriesDataType)reader.ReadUInt16();
            this.cValBSize = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
