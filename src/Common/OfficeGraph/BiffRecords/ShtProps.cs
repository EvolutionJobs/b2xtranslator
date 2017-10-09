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

namespace DIaLOGIKa.b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies properties of a chart as defined by the Chart Sheet Substream ABNF.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.ShtProps)]
    public class ShtProps : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.ShtProps;

        public enum EmptyCellPlotMode
        {
            PlotNothing,
            PlotAsZero,
            PlotAsInterpolated
        }

        /// <summary>
        /// A bit that specifies whether series are automatically allocated for the chart.
        /// </summary>
        public bool fManSerAlloc;

        /// <summary>
        /// If fAlwaysAutoPlotArea is true then this field MUST be true. 
        /// If fAlwaysAutoPlotArea is false then this field MUST be ignored.
        /// </summary>
        public bool fManPlotArea;

        /// <summary>
        /// A bit that specifies whether the default plot area dimension is used.<br/> 
        /// MUST be a value from the following table:
        /// </summary>
        public bool fAlwaysAutoPlotArea;

        /// <summary>
        /// Specifies how the empty cells are plotted.
        /// </summary>
        public EmptyCellPlotMode mdBlank;

        public ShtProps(IStreamReader reader, GraphRecordNumber id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            UInt16 flags = reader.ReadUInt16();
            this.fManSerAlloc = Utils.BitmaskToBool(flags, 0x1);
            // 0x2 and 0x4 are reserved
            this.fManPlotArea = Utils.BitmaskToBool(flags, 0x8);
            this.fAlwaysAutoPlotArea = Utils.BitmaskToBool(flags, 0x10);
            this.mdBlank = (EmptyCellPlotMode)reader.ReadByte();
            reader.ReadByte(); // skip the last byte

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
