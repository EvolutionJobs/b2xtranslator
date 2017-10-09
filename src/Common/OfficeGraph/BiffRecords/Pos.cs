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
    /// This record specifies the size and position for a legend, an attached label, 
    /// or the plot area, as specified by the primary axis group. <br/>
    /// This record MUST be ignored for the plot area when the fManPlotArea field of ShtProps is set to 1.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Pos)]
    public class Pos : OfficeGraphBiffRecord
    {
        public enum PositionMode
        {
            MDFX,
            MDABS,
            MDPARENT,
            MDKTH,
            MDCHART
        }

        public const GraphRecordNumber ID = GraphRecordNumber.Pos;

        /// <summary>
        /// A PositionMode that specifies the positioning mode for the upper-left corner of a legend, 
        /// an attached label, or the plot area. 
        /// </summary>
        public PositionMode mdTopLt;

        /// <summary>
        /// A PositionMode that specifies the positioning mode for the lower-right corner of a legend, 
        /// an attached label, or the plot area.
        /// </summary>
        public PositionMode mdBotRt;

        /// <summary>
        /// A signed integer that specifies a position. <br/>
        /// The meaning is specified in the Valid Combinations of mdTopLt and mdBotRt by Type table.
        /// </summary>
        public Int16 x1;

        /// <summary>
        /// A signed integer that specifies a position. <br/>
        /// The meaning is specified in the Valid Combinations of mdTopLt and mdBotRt by Type table.
        /// </summary>
        public Int16 y1;

        /// <summary>
        /// A signed integer that specifies a position. <br/>
        /// The meaning is specified in the Valid Combinations of mdTopLt and mdBotRt by Type table.
        /// </summary>
        public Int16 x2;

        /// <summary>
        /// A signed integer that specifies a position. <br/>
        /// The meaning is specified in the Valid Combinations of mdTopLt and mdBotRt by Type table.
        /// </summary>
        public Int16 y2;

        public Pos(IStreamReader reader, GraphRecordNumber id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.mdTopLt = (PositionMode)reader.ReadInt16();
            this.mdBotRt = (PositionMode)reader.ReadInt16();
            this.x1 = reader.ReadInt16();
            reader.ReadBytes(2); // skip 2 bytes
            this.y1 = reader.ReadInt16();
            reader.ReadBytes(2); // skip 2 bytes
            this.x2 = reader.ReadInt16();
            reader.ReadBytes(2); // skip 2 bytes
            this.y2 = reader.ReadInt16();
            reader.ReadBytes(2); // skip 2 bytes

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
