/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
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
using System.Text;
using System.IO;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    //[OfficeRecordAttribute(XXXX)]
    public class Pictures : Record
    {
        public Dictionary<long, Record> _pictures = new Dictionary<long, Record>();

        public Pictures(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.Reader.BaseStream.Position = 0;
            long pos;
            while (this.Reader.BaseStream.Position < this.Reader.BaseStream.Length)
            {
                pos = this.Reader.BaseStream.Position;
                var r = Record.ReadRecord(this.Reader);
                switch (r.TypeCode)
                {
                    case 0:
                        this.Reader.BaseStream.Position = this.Reader.BaseStream.Length;
                        break;
                    case 0xF01A:
                    case 0xF01B:
                    case 0xF01C:
                        var mb = (MetafilePictBlip)r;
                        _pictures.Add(pos, mb);
                        break;
                    case 0xF01D:
                    case 0xF01E:
                    case 0xF01F:
                    case 0xF020:
                    case 0xF021:
                        var b = (BitmapBlip)r;
                        _pictures.Add(pos, b);
                        break;
                    default:
                        break;
                }
                
            }
            pos = 1;
        }
    }
}
