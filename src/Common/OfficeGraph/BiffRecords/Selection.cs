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

namespace DIaLOGIKa.b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the selection within the data sheet window.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Selection)]
    public class Selection : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Selection;

        /// <summary>
        /// An unsigned integer that MUST be 0x03.
        /// </summary>
        public byte pnn;

        /// <summary>
        /// Specifies the row number of the active cell.<br/>
        /// MUST be greater than or equal to rwFirst.<br/> 
        /// MUST be less than or equal to rwLast.
        /// </summary>
        public ushort rwAct;

        /// <summary>
        /// A Graph_Col that specifies the column number of the active cell.<br/>
        /// MUST be greater than or equal to colFirst.<br/> 
        /// MUST be less than or equal to colLast.
        /// </summary>
        public ushort colAct;

        /// <summary>
        /// Specifies the topmost row of the active selection.
        /// </summary>
        public ushort rwFirst;

        /// <summary>
        /// Specifies bottommost row of the active selection.
        /// </summary>
        public ushort rwLast;

        /// <summary>
        /// Specifies the leftmost column of the active selection.
        /// </summary>
        public ushort colFirst;

        /// <summary>
        /// Specifies the rightmost column of the active selection.
        /// </summary>
        public ushort colLast;

        public Selection(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.pnn = reader.ReadByte();
            this.rwAct = reader.ReadUInt16();
            this.colAct = reader.ReadUInt16();
            reader.ReadBytes(4); //skip 4 bytes
            this.rwFirst = reader.ReadUInt16();
            this.rwLast = reader.ReadUInt16();
            this.colFirst = reader.ReadUInt16();
            this.colLast = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
