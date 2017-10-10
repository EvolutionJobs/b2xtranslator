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
using System.IO;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies the properties of a fill pattern for parts of a chart. 
    /// The record consists of an OfficeArtFOPT, as specified in [MS-ODRAW] section 2.2.9, 
    /// and an OfficeArtTertiaryFOPT, as specified in [MS-ODRAW] section 2.2.11, that both
    /// contain properties for the fill pattern applied. <55>
    /// </summary>
    [BiffRecordAttribute(RecordType.GelFrame)]
    public class GelFrame : BiffRecord
    {
        public const RecordType ID = RecordType.GelFrame;

        public Record OPT1;

        public Record OPT2;

        public GelFrame(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // copy data to a memory stream to cope with Continue records
            using (var ms = new MemoryStream())
            {
                var buffer = reader.ReadBytes(length);
                ms.Write(buffer, 0, length);

                if (BiffRecord.GetNextRecordType(reader) == RecordType.GelFrame)
                {
                    var nextId = (RecordType)reader.ReadUInt16();
                    var nextLength = reader.ReadUInt16();

                    buffer = reader.ReadBytes(nextLength);
                    ms.Write(buffer, 0, nextLength);
                }

                while (BiffRecord.GetNextRecordType(reader) == RecordType.Continue)
                {
                    var nextId = (RecordType)reader.ReadUInt16();
                    var nextLength = reader.ReadUInt16();

                    buffer = reader.ReadBytes(nextLength);
                    ms.Write(buffer, 0, nextLength);
                }

                ms.Position = 0;

                // initialize class members from stream
                this.OPT1 = Record.ReadRecord(ms);

                if (ms.Position < ms.Length)
                {
                    this.OPT2 = Record.ReadRecord(ms);
                }
            }
            
            // assert that the correct number of bytes has been read from the stream
            //Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
