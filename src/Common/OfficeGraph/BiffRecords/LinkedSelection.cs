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
    /// This record specifies where in the data sheet window to paste the selection from the OLE stream.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.LinkedSelection)]
    public class LinkedSelection : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.LinkedSelection;

        /// <summary>
        /// Specifies the first row in the data sheet window in which to paste the selection from the OLE stream.<br/>
        /// MUST be 0x0000 if the first row of the selection from the OLE stream contains any non-numeric values.<br/> 
        /// MUST be 0x0001 if the first row of the selection from the OLE stream contains only numeric values.
        /// </summary>
        public ushort rwFirst;

        /// <summary>
        /// MUST be the same as rwFirst.
        /// </summary>
        public ushort rwLast;

        /// <summary>
        /// Specifies the first column in the data sheet window in which to paste the selection from the OLE stream. <br/>
        /// MUST be 0x0000 if the first column of the selection from the OLE stream contains any non-numeric values. <br/>
        /// MUST be 0x0001 if the first column of the selection from the OLE stream contains only numeric values.
        /// </summary>
        public ushort colFirst;

        /// <summary>
        /// MUST be the same as colFirst.
        /// </summary>
        public ushort colLast;

        public LinkedSelection(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.rwFirst = reader.ReadUInt16();
            this.rwLast = reader.ReadUInt16();
            this.colFirst = reader.ReadUInt16();
            this.colLast = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
