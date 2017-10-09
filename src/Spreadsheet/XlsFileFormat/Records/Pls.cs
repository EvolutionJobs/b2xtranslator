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
using System.IO;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecordAttribute(RecordType.Pls)] 
    public class Pls : BiffRecord
    {
        public const RecordType ID = RecordType.Pls;

        public byte[] rgb;

        public Pls(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            using (MemoryStream ms = new MemoryStream())
            {
                byte[] buffer = reader.ReadBytes(length);
                ms.Write(buffer, 0, length);

                while (BiffRecord.GetNextRecordType(reader) == RecordType.Pls)
                {
                    RecordType nextId = (RecordType)reader.ReadUInt16();
                    UInt16 nextLength = reader.ReadUInt16();

                    buffer = reader.ReadBytes(nextLength);
                    ms.Write(buffer, 0, nextLength);
                }

                while (BiffRecord.GetNextRecordType(reader) == RecordType.Continue)
                {
                    RecordType nextId = (RecordType)reader.ReadUInt16();
                    UInt16 nextLength = reader.ReadUInt16();

                    buffer = reader.ReadBytes(nextLength);
                    ms.Write(buffer, 0, nextLength);
                }

                ms.Position = 0;

                // initialize class members from stream
                this.rgb = new byte[ms.Length];
                ms.Read(this.rgb, 0, (int)ms.Length);

                // assert that the correct number of bytes has been read from the stream
                //Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
            }
        }
    }
}
