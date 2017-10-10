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
    /// This record specifies the number of non-empty rows and the number of non-empty cells in the longest row of a Graph object.
    /// </summary>
    [BiffRecordAttribute(RecordType.Dimensions)]
    public class Dimensions : BiffRecord
    {
        public const RecordType ID = RecordType.Dimensions;

        /// <summary>
        /// A RwLongU that specifies the first row in the sheet that contains a used cell.
        /// </summary>
        public uint rwMic;

        /// <summary>
        /// An unsigned integer that specifies the number of non-empty cells in the 
        /// longest row in the data sheet of a Graph object. 
        /// 
        /// MUST be less than or equal to 0x00000F9F.
        /// </summary>
        public uint rwMac;

        /// <summary>
        /// A ColU that specifies the first column in the sheet that contains a used cell.
        /// </summary>
        public ushort colMic;

        /// <summary>
        /// An unsigned integer that specifies the number of non-empty rows in the 
        /// data sheet of a Graph object. 
        /// 
        /// MUST be less than or equal to 0x00FF.
        /// </summary>
        public ushort colMac;

        public Dimensions(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.rwMic = reader.ReadUInt32();
            this.rwMac = reader.ReadUInt32();
            this.colMic = reader.ReadUInt16();
            this.colMac = reader.ReadUInt16();
            reader.ReadBytes(2);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
