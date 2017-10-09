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
using System.Collections.Generic;
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecordAttribute(RecordType.CRN)] 
    public class CRN : BiffRecord
    {
        public const RecordType ID = RecordType.CRN;

        public byte colLast;
        public byte colFirst;

        public UInt16 rw;

        public List<Object> oper; 


        public CRN(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);
            this.oper = new List<object>(); 
            long endposition = this.Reader.BaseStream.Position + this.Length; 
            this.colLast = this.Reader.ReadByte();
            this.colFirst = this.Reader.ReadByte();
            this.rw = this.Reader.ReadUInt16();

            


            while (this.Reader.BaseStream.Position < endposition)
            {
                byte grbit = this.Reader.ReadByte();
                if (grbit == 0x01)
                {
                    this.oper.Add(this.Reader.ReadDouble()); 
                }
                else if (grbit == 0x02)
                {
                    String Data = "";
                    UInt16 cch = this.Reader.ReadUInt16();
                    byte firstbyte = this.Reader.ReadByte();
                    int firstbit = firstbyte & 0x1;
                    for (int i = 0; i < cch; i++)
                    {
                        if (firstbit == 0)
                        {
                            Data += (char)this.Reader.ReadByte();
                            // read 1 byte per char 
                        }
                        else
                        {
                            // read two byte per char 
                            Data += System.BitConverter.ToChar(this.Reader.ReadBytes(2), 0);
                        }
                    }

                    this.oper.Add(Data); 
                }
                else if (grbit == 0x00)
                {
                    this.Reader.ReadBytes(8);
                    this.oper.Add(" "); 
                }
                else if (grbit == 0x04)
                {
                    // bool 
                    UInt16 boolvalue = this.Reader.ReadUInt16();
                    bool value = false;
                    if (boolvalue == 1)
                        value = true;
                    this.oper.Add(value);
                    this.Reader.ReadBytes(6); 
                }
                else if (grbit == 0x10)
                {
                    // Error
                    this.Reader.ReadBytes(8);
                    this.oper.Add("Err"); 
                }

            }
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
