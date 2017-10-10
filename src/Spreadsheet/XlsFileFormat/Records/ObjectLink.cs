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

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies an object on a chart, or the entire chart, to which Text is linked.
    /// </summary>
    [BiffRecordAttribute(RecordType.ObjectLink)]
    public class ObjectLink : BiffRecord
    {
        public const RecordType ID = RecordType.ObjectLink;

        public enum ObjectType : ushort
        {
            Chart = 0x0001,
            DVAxis = 0x0002,
            IVAxis = 0x0003,
            SeriesOrDataPoints = 0x0004,
            SeriesAxis = 0x0007,
            DisplayUnitsLabels = 0x000C
        }

        /// <summary>
        /// An unsigned integer that specifies the object that the Text is linked to. <br/>
        /// MUST be a value from the following table:<br/>
        /// 0x0001 = Entire chart.
        /// 0x0002 = Value axis, or vertical value axis on bubble and scatter chart groups.
        /// 0x0003 = category (3) axis, or horizontal value axis on bubble and scatter chart groups.
        /// 0x0004 = Series or data points.
        /// 0x0007 = Series axis.
        /// 0x000C = Display units labels of an axis.
        /// </summary>
        public ObjectType wLinkObj;

        /// <summary>
        /// An unsigned integer that specifies the zero-based index into a Series record in the collection 
        /// of Series records in the current chart sheet substream. 
        /// Each referenced Series record specifies a series for the chart group to which the Text is linked. 
        /// When the wLinkObj field is 4, MUST be less than or equal to 254. 
        /// When the wLinkObj field is not 4, MUST be zero, and MUST be ignored.
        /// </summary>
        public ushort wLinkVar1;

        /// <summary>
        /// An unsigned integer that specifies the zero-based index into the category (3) within the series 
        /// specified by wLinkVar1, to which the Text is linked. 
        /// When the wLinkObj field is 4, if the Text is linked to a series instead of a single data point, 
        /// the value MUST be 0xFFFF; if the Text is linked to a data point, 
        /// the value MUST be less than or equal to 31999. 
        /// When the wLinkObj field is not 4, MUST be zero, and MUST be ignored.
        /// </summary>
        public ushort wLinkVar2;

        public ObjectLink(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.wLinkObj = (ObjectType)reader.ReadUInt16();
            this.wLinkVar1 = reader.ReadUInt16();
            this.wLinkVar2 = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
