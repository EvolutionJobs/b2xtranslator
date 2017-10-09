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
using System.Collections;
using DIaLOGIKa.b2xtranslator.Tools;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class FontFamilyName : ByteStructure
    {
        public struct FontSignature
        {
            public UInt32 UnicodeSubsetBitfield0;
            public UInt32 UnicodeSubsetBitfield1;
            public UInt32 UnicodeSubsetBitfield2;
            public UInt32 UnicodeSubsetBitfield3;
            public UInt32 CodePageBitfield0;
            public UInt32 CodePageBitfield1;
        }

        /// <summary>
        /// When true, font is a TrueType font
        /// </summary>
        public bool fTrueType;

        /// <summary>
        /// Font family id
        /// </summary>
        public byte ff;

        /// <summary>
        /// Base weight of font
        /// </summary>
        public Int16 wWeight;

        /// <summary>
        /// Character set identifier
        /// </summary>
        public byte chs;

        /// <summary>
        /// Pitch request
        /// </summary>
        public byte prq;

        /// <summary>
        /// Name of font
        /// </summary>
        public String xszFtn;

        /// <summary>
        /// Alternative name of the font
        /// </summary>
        public String xszAlt;

        /// <summary>
        /// Panose
        /// </summary>
        public byte[] panose;

        /// <summary>
        /// Font sinature
        /// </summary>
        public FontSignature fs;


        public FontFamilyName(VirtualStreamReader reader, int length) : base(reader, length)
        {
            long startPos = _reader.BaseStream.Position;

            //FFID
            int ffid = (int)_reader.ReadByte();

            int req = ffid;
            req = req << 6;
            req = req >> 6;
            this.prq = (byte)req;

            this.fTrueType = Utils.BitmaskToBool(ffid, 0x04);

            int family = ffid;
            family = family << 1;
            family = family >> 4;
            this.ff = (byte)family;

            this.wWeight = _reader.ReadInt16();

            this.chs = _reader.ReadByte();

            //skip byte 5
            _reader.ReadByte();

            //read the 10 bytes panose
            this.panose = _reader.ReadBytes(10);

            //read the 24 bytes FontSignature
            this.fs = new FontSignature();
            this.fs.UnicodeSubsetBitfield0 = _reader.ReadUInt32();
            this.fs.UnicodeSubsetBitfield1 = _reader.ReadUInt32();
            this.fs.UnicodeSubsetBitfield2 = _reader.ReadUInt32();
            this.fs.UnicodeSubsetBitfield3 = _reader.ReadUInt32();
            this.fs.CodePageBitfield0 = _reader.ReadUInt32();
            this.fs.CodePageBitfield1 = _reader.ReadUInt32();

            //read the next \0 terminated string
            long strStart = reader.BaseStream.Position;
            long strEnd = searchTerminationZero(_reader);
            this.xszFtn = Encoding.Unicode.GetString(_reader.ReadBytes((int)(strEnd - strStart)));
            this.xszFtn = this.xszFtn.Replace("\0", "");

            long readBytes = _reader.BaseStream.Position - startPos;
            if(readBytes < _length)
            {
                //read the next \0 terminated string
                strStart = reader.BaseStream.Position;
                strEnd = searchTerminationZero(_reader);
                this.xszAlt = Encoding.Unicode.GetString(_reader.ReadBytes((int)(strEnd - strStart)));
                this.xszAlt = this.xszAlt.Replace("\0", "");
            }
        }

        private long searchTerminationZero(VirtualStreamReader reader)
        {
            long strStart = reader.BaseStream.Position;
            while (reader.ReadInt16() != 0)
            {
                ;
            }
            long pos = reader.BaseStream.Position;
            reader.BaseStream.Seek(strStart, System.IO.SeekOrigin.Begin);
            return pos;
        }
    }
}
