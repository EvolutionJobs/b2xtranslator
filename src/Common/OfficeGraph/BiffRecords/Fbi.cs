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

namespace DIaLOGIKa.b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the font information at the time the scalable font is added to the chart. <47>
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Fbi)]
    public class Fbi : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Fbi;

        /// <summary>
        /// An unsigned integer that specifies the font width, in twips, when the font was first applied. 
        /// 
        /// MUST be greater than or equal to 0 and less than or equal to 0x7FFF.
        /// </summary>
        public ushort dmixBasis;

        /// <summary>
        /// An unsigned integer that specifies the font height, in twips, when the font was first applied. 
        /// 
        /// MUST be greater than or equal to 0 and less than or equal to 0x7FFF.
        /// </summary>
        public ushort dmiyBasis;

        /// <summary>
        /// An unsigned integer that specifies the default font height in twips. 
        /// 
        /// MUST be greater than or equal to 20 and less than or equal to 8180.
        /// </summary>
        public ushort twpHeightBasis;

        /// <summary>
        /// A Boolean that specifies the scale to use. The value MUST be one of the following values: 
        /// 
        ///     Value       Meaning
        ///     0x0000      Scale by chart area
        ///     0x0001      Scale by plot area
        /// </summary>
        public bool scab;

        public ushort ifnt;
        // TODO: implement FontIndex???

        public Fbi(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.dmixBasis = reader.ReadUInt16();
            this.dmiyBasis = reader.ReadUInt16();
            this.twpHeightBasis = reader.ReadUInt16();
            this.scab = reader.ReadUInt16() == 0x0001;
            this.ifnt = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
