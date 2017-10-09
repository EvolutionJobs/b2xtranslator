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
    /// This record specifies that the chart group is a pie chart group or a doughnut chart group, 
    /// and specifies the chart group attributes.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Pie)]
    public class Pie : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Pie;

        /// <summary>
        /// An unsigned integer that specifies the starting angle of the first data point, 
        /// clockwise from the top of the circle. <br/>
        /// MUST be less than or equal to 360.
        /// </summary>
        public UInt16 anStart;

        /// <summary>
        /// An unsigned integer that specifies the size of the center hole in a doughnut chart group as a 
        /// percentage of the plot area size. <br/>
        /// MUST be a value from the following table:<br/>
        /// 0 = Pie chart group<br/>
        /// 10 to 90 =  Doughnut chart group
        /// </summary>
        public UInt16 pcDonut;

        /// <summary>
        /// A bit that specifies whether one or more data points in the chart group has shadows.
        /// </summary>
        public bool fHasShadow;

        /// <summary>
        /// A bit that specifies whether the leader lines to the data labels are shown.
        /// </summary>
        public bool fShowLdrLines;

        public Pie(IStreamReader reader, GraphRecordNumber id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.anStart = reader.ReadUInt16();
            this.pcDonut = reader.ReadUInt16();
            UInt16 flags = reader.ReadUInt16();
            this.fHasShadow = Utils.BitmaskToBool(flags, 0x1);
            this.fShowLdrLines = Utils.BitmaskToBool(flags, 0x2);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
