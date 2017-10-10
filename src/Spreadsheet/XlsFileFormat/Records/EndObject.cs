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
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies properties of a Future Record Type (FRT) as defined 
    /// by the Chart Sheet Substream ABNF. The collection of records specifies a 
    /// feature saved as an FRT such that an application not supporting 
    /// the feature can preserve it. 
    /// 
    /// This record MUST have an associated StartObject record. StartObject and 
    /// EndObject pairs can be nested. Up to 100 levels of blocks can be nested.
    /// </summary>
    [BiffRecordAttribute(RecordType.EndObject)]
    public class EndObject : BiffRecord
    {
        public const RecordType ID = RecordType.EndObject;

        public enum ObjectKind : ushort
        {
            YMult = 0x0010,
            FrtFontList = 0x0011,
            DataLabExt = 0x0012
        }

        /// <summary>
        /// An FrtHeaderOld. The frtHeaderOld.rt field MUST be 0x0855.
        /// </summary>
        public FrtHeaderOld frtHeaderOld;

        /// <summary>
        /// An unsigned integer that specifies the type of object that is encompassed by the block. 
        /// 
        /// MUST equal the iObjectKind field of the associated StartObject record.
        /// </summary>
        public ObjectKind iObjectKind;


        public EndObject(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.frtHeaderOld = new FrtHeaderOld(reader);
            this.iObjectKind = (ObjectKind)reader.ReadUInt16();

            reader.ReadBytes(6);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
