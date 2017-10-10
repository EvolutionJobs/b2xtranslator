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
    /// This record specifies a number format.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Format)]
    public class Format : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Format;

        /// <summary>
        /// An IFmt that specifies the identifier of the format string specified by stFormat. 
        /// 
        /// The value of ifmt.ifmt SHOULD <52> be a value within one of the following ranges. 
        /// The value of ifmt.ifmt MUST be a value within one of the following ranges or within 383 to 392.
        ///     -   5 to 8
        ///     -  23 to 26
        ///     -  41 to 44
        ///     -  63 to 66
        ///     - 164 to 382
        /// </summary>
        public ushort ifmt;

        /// <summary>
        /// An XLUnicodeString that specifies the format string for this number format. 
        /// The format string indicates how to format the numeric value of the cell. 
        /// 
        /// The length of this field MUST be greater than or equal to 1 character and less than 
        /// or equal to 255 characters. For more information about how format strings are 
        /// interpreted, see [ECMA-376] Part 4: Markup Language Reference, section 3.8.31.
        /// 
        /// The ABNF grammar for the format string is specified in [MS-XLS] section 2.4.126.
        /// </summary>
        public XLUnicodeString stFormat;

        public Format(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.ifmt = reader.ReadUInt16();
            this.stFormat = new XLUnicodeString(reader);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
