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

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies the beginning of a collection of records as defined by 
    /// the Chart Sheet Substream ABNF. The collection of records specifies the options 
    /// of the data table which can be displayed within a chart area.
    /// </summary>
    [BiffRecordAttribute(RecordType.Dat)]
    public class Dat : BiffRecord
    {
        public const RecordType ID = RecordType.Dat;

        /// <summary>
        /// A bit that specifies whether horizontal cell borders are displayed within the data table.
        /// </summary>
        public bool fHasBordHorz;

        /// <summary>
        /// A bit that specifies whether vertical cell borders are displayed within the data table.
        /// </summary>
        public bool fHasBordVert;

        /// <summary>
        /// A bit that specifies whether an outside outline is displayed around the data table.
        /// </summary>
        public bool fHasBordOutline;

        /// <summary>
        /// A bit that specifies whether the legend key is displayed next to the name of the series. 
        /// If the value is 1, the legend key symbols are displayed next to the name of the series.
        /// </summary>
        public bool fShowSeriesKey;

        public Dat(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            var flags = reader.ReadUInt16();

            this.fHasBordHorz = Utils.BitmaskToBool(flags, 0x0001);
            this.fHasBordVert = Utils.BitmaskToBool(flags, 0x0002);
            this.fHasBordOutline = Utils.BitmaskToBool(flags, 0x0004);
            this.fShowSeriesKey = Utils.BitmaskToBool(flags, 0x0008);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
