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
    /// This record specifies the size and position of the data sheet window within the 
    /// OLE server window that is contained in the parent document window. 
    /// 
    /// MUST immediately follow a MainWindow record.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Window1_10)]
    public class Window1_10 : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Window1_10;

        /// <summary>
        /// An unsigned integer that specifies the X location of the upper-left 
        /// corner of the data sheet window within the OLE server window, in twips.
        /// </summary>
        public ushort xWn;

        /// <summary>
        /// An unsigned integer that specifies the Y location of the upper-left 
        /// corner of the data sheet window within the OLE server window, in twips.
        /// </summary>
        public ushort yWn;

        /// <summary>
        /// An unsigned integer that specifies the width of the data sheet window 
        /// within the OLE server window, in twips. MUST be greater than or equal to 0x0001.
        /// </summary>
        public ushort dxWn;

        /// <summary>
        /// An unsigned integer that specifies the height of the data sheet window 
        /// within the OLE server window, in twips. MUST be greater than or equal to 0x0001.
        /// </summary>
        public ushort dyWn;

        public Window1_10(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.xWn = reader.ReadUInt16();
            this.yWn = reader.ReadUInt16();
            this.dxWn = reader.ReadUInt16();
            this.dyWn = reader.ReadUInt16();

            // skipped reserved bytes
            reader.ReadBytes(2);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
