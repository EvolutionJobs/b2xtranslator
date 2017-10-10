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
    [BiffRecordAttribute(RecordType.ColInfo)] 
    public class ColInfo : BiffRecord
    {
        public const RecordType ID = RecordType.ColInfo;

        public int colFirst;
        public int colLast;

        public int coldx;

        public bool fUserSet;
        public bool fHidden;
        public bool fBestFit;
        public bool fPhonetic;
        public int iOutLevel;
        public bool fCollapsed;

        public int ixfe;

        public ColInfo(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.colFirst = reader.ReadUInt16();
            this.colLast = reader.ReadUInt16();
            this.coldx = reader.ReadUInt16();
            this.ixfe = reader.ReadUInt16();

            int buffer = reader.ReadUInt16(); 
            
            ///
            /// A - fHidden (1 bit)
            /// B - fUserSet (1 bit)
            /// C - fBestFit (1 bit)
            /// D - fPhonetic (1 bit)
            /// E - reserved1 (4 bits): MUST be zero, and MUST be ignored.
            /// F - iOutLevel (3 bits)
            /// G - unused1 (1 bit): Undefined and MUST be ignored.
            /// H - fCollapsed (1 bit)
            /// I - reserved2 (3 bits): MUST be zero, and MUST be ignored.
            this.fHidden = Utils.BitmaskToBool(buffer, 0x0001);
            this.fUserSet = Utils.BitmaskToBool(buffer, 0x0002);
            this.fBestFit = Utils.BitmaskToBool(buffer, 0x0004);
            this.fPhonetic = Utils.BitmaskToBool(buffer, 0x0008);

            this.iOutLevel = (int)(buffer & 0x0700) >> 0x8;

            this.fCollapsed = Utils.BitmaskToBool(buffer, 0x1000); 

            // read two following not documented bytes 
            reader.ReadUInt16(); 

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
