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
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.IO;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(4009)]
    public class TextSIExceptionAtom : Record
    {
        public TextSIException si;
        public TextSIExceptionAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            si = new TextSIException(Reader);
        }
    }

    [OfficeRecordAttribute(4010)]
    public class TextSpecialInfoAtom : Record
    {
        public List<TextSIRun> Runs = new List<TextSIRun>();

        public TextSpecialInfoAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

            while (Reader.BaseStream.Position < Reader.BaseStream.Length)
            {
                TextSIRun run = new TextSIRun(Reader);
                Runs.Add(run);
            }

        }       
    }

    public class TextSIRun
    {
        public uint count;
        public TextSIException si;

        public TextSIRun(BinaryReader reader)
        {
            count = reader.ReadUInt32();
            si = new TextSIException(reader);
        }
    }

    public class TextSIException
    {
        private uint flags;
        public bool spell;
        public bool lang;
        public bool altLang;
        public bool fPp10ext;
        public bool fBidi;
        public bool smartTag;
        public UInt16 spellInfo;
        public UInt16 lid;
        public UInt16 bidi;
        public UInt16 altLid;
        
        public TextSIException(BinaryReader reader)
        {
            flags = reader.ReadUInt32();

            spell = Utils.BitmaskToBool(flags, 0x1);
            lang = Utils.BitmaskToBool(flags, 0x1 << 1);
            altLang = Utils.BitmaskToBool(flags, 0x1 << 2);
            fPp10ext = Utils.BitmaskToBool(flags, 0x1 << 5);
            fBidi = Utils.BitmaskToBool(flags, 0x1 << 6);
            smartTag = Utils.BitmaskToBool(flags, 0x1 << 9);

            if (spell) spellInfo = reader.ReadUInt16();
            if (lang) lid = reader.ReadUInt16();
            if (altLang) altLid = reader.ReadUInt16();
            if (fBidi) bidi = reader.ReadUInt16();
            UInt32 dummy;
            if (fPp10ext) dummy = reader.ReadUInt32();
            byte[] smartTags;
            if (smartTag) smartTags = reader.ReadBytes((int)(reader.BaseStream.Length - reader.BaseStream.Position));

        }
    }
}
