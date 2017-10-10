/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */
using System;
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecordAttribute(RecordType.ShapePropsStream)]
    public class ShapePropsStream : BiffRecord
    {
        public const RecordType ID = RecordType.ShapePropsStream;

        public FrtHeader frtHeader;

        /// <summary>
        /// An unsigned integer that specifies the chart element that the shape 
        /// formatting properties in this record apply to.
        /// </summary>
        public ushort wObjContext;

        /// <summary>
        /// An unsigned integer that specifies the checksum of the shape formatting properties related to this record.
        /// </summary>
        public UInt32 dwChecksum;

        /// <summary>
        /// An unsigned integer that specifies the length of the character array in the rgb field.
        /// </summary>
        public UInt32 cb;

        /// <summary>
        /// An array of ANSI characters, whose length is specified by cb, 
        /// that contains the XML representation of the shape formatting properties 
        /// as defined in [ECMA-376] Part 4, section 5.7.2.198.
        /// </summary>
        public byte[] rgb;

        public ShapePropsStream(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.frtHeader = new FrtHeader(reader);
            this.wObjContext = reader.ReadUInt16();
            reader.ReadBytes(2); //unused
            this.dwChecksum = reader.ReadUInt32();
            this.cb = reader.ReadUInt32();
            this.rgb = reader.ReadBytes((int)cb);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
