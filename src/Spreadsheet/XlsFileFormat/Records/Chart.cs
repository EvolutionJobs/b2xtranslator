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
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies the position and size of the chart area and specifies the beginning 
    /// of a collection of records as defined by the Chart Sheet Substream ABNF. The collection of records specifies a chart.
    /// </summary>
    [BiffRecordAttribute(RecordType.Chart)]
    public class Chart : BiffRecord
    {
        public const RecordType ID = RecordType.Chart;

        /// <summary>
        /// A FixedPoint as specified in [MS-OSHARED] section 2.2.1.6 that specifies the 
        /// horizontal position of the upper-left corner of the chart in points. 
        /// 
        /// SHOULD <40> be greater than or equal to zero.
        /// </summary>
        public FixedPointNumber x;

        /// <summary>
        /// A FixedPoint as specified in [MS-OSHARED] section 2.2.1.6 that specifies the 
        /// horizontal position of the upper-left corner of the chart in points. 
        /// 
        /// SHOULD <40> be greater than or equal to zero.
        /// </summary>
        public FixedPointNumber y;

        /// <summary>
        /// A FixedPoint as specified in [MS-OSHARED] section 2.2.1.6 that specifies 
        /// the width in points.
        /// 
        /// MUST be greater than or equal to 0.
        /// </summary>
        public FixedPointNumber dx;

        /// <summary>
        /// A FixedPoint as specified in [MS-OSHARED] section 2.2.1.6 that specifies
        /// the height in points. 
        /// 
        /// MUST be greater than or equal to 0.
        /// </summary>
        public FixedPointNumber dy;

        public Chart(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.x = new FixedPointNumber(reader);
            this.y = new FixedPointNumber(reader);
            this.dx = new FixedPointNumber(reader);
            this.dy = new FixedPointNumber(reader);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
