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
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies how the series data of a chart is arranged.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Orient)]
    public class Orient : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Orient;

        /// <summary>
        /// A Boolean that specifies whether series are arranged by rows or columns from the 
        /// data specified in a data sheet window. <br/>
        /// MUST be a value from the following table:<br/>
        /// false = Series are arranged by columns.<br/>
        /// true = Series are arranged by rows.
        /// </summary>
        public bool fSeriesInRows;

        /// <summary>
        /// An unsigned integer that specifies a zero-based row index into the data sheet. <br/>
        /// The referenced row is used to calculate axis values along the horizontal value axis in a 
        /// scatter chart group. MUST be 0x0000 for other chart group types. <br/>
        /// MUST equal to colSeriesX and MUST be ignored if fSeriesInRows is 0x00.
        /// </summary>
        public ushort rowSeriesX;

        /// <summary>
        /// An unsigned integer that specifies a zero-based column index into the data sheet.<br/>
        /// The referenced column is used to calculate axis values along the horizontal value 
        /// axis in a scatter chart group. MUST be 0x0000 for other chart group types.<br/> 
        /// MUST equal to rowSeriesX and MUST be ignored if fSeriesInRows is 0x01.
        /// </summary>
        public ushort colSeriesX;

        public Orient(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.fSeriesInRows = Utils.ByteToBool(reader.ReadByte());
            this.rowSeriesX = reader.ReadUInt16();
            this.colSeriesX = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
