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
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecordAttribute(RecordType.ExternName)] 
    public class ExternName : BiffRecord
    {
        public const RecordType ID = RecordType.ExternName;

        public ushort ixals;
        public bool fOle;
        public bool fOleLink;
        public ushort grbit; 
        public string extName;
        public byte cch;

        public ushort cce;

        public string nameDefinition; 

        public ExternName(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.grbit = this.Reader.ReadUInt16();



            this.fOle = Utils.BitmaskToBool(this.grbit, 0x0008);
            this.fOleLink = Utils.BitmaskToBool(this.grbit, 0x0010);

            // this is an external link 
            if (!this.fOle && !this.fOleLink)
            {

                this.ixals = this.Reader.ReadUInt16();
                // read unused 16 bit
                this.Reader.ReadBytes(2);
                this.cch = this.Reader.ReadByte();
                byte firstbyte = this.Reader.ReadByte();
                int firstbit = firstbyte & 0x1;
                for (int i = 0; i < this.cch; i++)
                {
                    if (firstbit == 0)
                    {
                        this.extName += (char)this.Reader.ReadByte();
                        // read 1 byte per char 
                    }
                    else
                    {
                        // read two byte per char 
                        this.extName += System.BitConverter.ToChar(this.Reader.ReadBytes(2), 0);
                    }
                }
                this.cce = this.Reader.ReadUInt16();
                this.Reader.ReadBytes(this.cce); 

            }
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
