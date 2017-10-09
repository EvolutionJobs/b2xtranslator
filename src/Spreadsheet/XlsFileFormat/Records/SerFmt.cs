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

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies properties of the associated data points, data markers, or lines of the series. 
    /// The associated data points, data markers, or lines of the series are specified by the preceding 
    /// DataFormat record. If this record is not present in the sequence of records that conforms to the 
    /// SS rule then all the fields will have default values otherwise all the fields MUST contain a value.
    /// </summary>
    [BiffRecordAttribute(RecordType.SerFmt)]
    public class SerFmt : BiffRecord
    {
        public const RecordType ID = RecordType.SerFmt;

        /// <summary>
        /// A bit that specifies whether the lines of the series are displayed with a 
        /// smooth line effect on a scatter, radar, and line chart group. <br/>
        /// The default value of this field is 0.
        /// </summary>
        public bool fSmoothedLine;

        /// <summary>
        /// A bit that specifies whether the data points of a bubble chart group are 
        /// displayed with a 3-D effect. <br/>
        /// MUST be ignored for all other chart groups. <br/>
        /// The default value of this field is 0.
        /// </summary>
        public bool f3DBubbles;

        /// <summary>
        /// A bit that specifies whether the data markers are displayed with a 
        /// shadow on bubble, scatter, radar, stock, and line chart groups.<br/> 
        /// The default value of this field is 0.
        /// </summary>
        public bool fArShadow;

        public SerFmt(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            UInt16 flags = reader.ReadUInt16();
            this.fSmoothedLine = Utils.BitmaskToBool(flags, 0x1);
            this.f3DBubbles = Utils.BitmaskToBool(flags, 0x2);
            this.fArShadow = Utils.BitmaskToBool(flags, 0x4);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
