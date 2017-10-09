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
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.PieFormat)]
    public class PieFormat : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.PieFormat;

        /// <summary>
        /// A signed integer that specifies the distance of a 
        /// data point or data points in a series from the center of one of the following:<br/>
        /// <list>
        /// <item>The plot area for a doughnut or pie chart group.</item>
        /// <item>The primary pie in a pie of pie or bar of pie chart group.</item>
        /// <item>The secondary bar/pie of a pie of pie chart group.</item>
        /// </list>
        /// The value of this field specifies the distance as a percentage. <br/>
        /// If this value is 0, then the data point or data points in a series is as close to 
        /// the center as possible for the particular chart group type. <br/>
        /// If this value is 100, then the data point is at the edge of the chart area. <br/>
        /// If this value is greater than 100, such that the data point is beyond the edge of the chart area, 
        /// then all the data points in the chart group are scaled down to fit inside the chart 
        /// area such that the data point with the highest pcExplode value is at the edge of the chart area.
        /// </summary>
        public Int16 pcExplode;

        public PieFormat(IStreamReader reader, GraphRecordNumber id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.pcExplode = reader.ReadInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
