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
    /// This record specifies an empty cell with no formula or value.
    /// </summary>
    [BiffRecordAttribute(RecordType.BlankGraph)]
    public class BlankGraph : BiffRecord
    {
        public const RecordType ID = RecordType.BlankGraph;

        /// <summary>
        /// An unsigned integer that specifies a zero-based index of a row in the datasheet that contains this structure. 
        /// 
        /// MUST be less than or equal to 0x0F9F.
        /// </summary>
        public ushort rw;

        /// <summary>
        /// An unsigned integer that specifies a zero-based index of a column in the datasheet that contains this structure. 
        /// 
        /// MUST be less than or equal to 0x0F9F.
        /// </summary>
        public ushort col;

        /// <summary>
        /// An unsigned integer that specifies the identifier of a number format. 
        /// 
        /// The identifier specified by this field MUST be a valid built-in number format identifier 
        /// or the identifier of a custom number format as specified using a Format record. 
        /// 
        /// Custom number format identifiers MUST be greater than or equal to 0x00A4 less than or equal to 0x0188, 
        /// and SHOULD be less than or equal to 0x017E. 
        /// 
        /// The built-in number formats are listed in [ECMA-376] Part 4: Markup Language Reference, section 3.8.30.
        /// </summary>
        public ushort ifmt;

        public BlankGraph(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.rw = reader.ReadUInt16();
            this.col = reader.ReadUInt16();

            // ignore reserved byte
            reader.ReadBytes(1);

            this.ifmt = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
