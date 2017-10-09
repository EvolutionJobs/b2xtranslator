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

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(2032)]
    public class ColorSchemeAtom : Record
    {
        public byte[] Bytes;
        public Int32 Background;
        public Int32 TextAndLines;
        public Int32 Shadows;
        public Int32 TitleText;
        public Int32 Fills;
        public Int32 Accent;
        public Int32 AccentAndHyperlink;
        public Int32 AccentAndFollowedHyperlink;

        public ColorSchemeAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {

                this.Bytes = this.Reader.ReadBytes((int)this.BodySize);
                this.Background = System.BitConverter.ToInt32(this.Bytes, 0);
                this.TextAndLines = System.BitConverter.ToInt32(this.Bytes, 4);
                this.Shadows = System.BitConverter.ToInt32(this.Bytes, 8);
                this.TitleText = System.BitConverter.ToInt32(this.Bytes, 12);
                this.Fills = System.BitConverter.ToInt32(this.Bytes, 16);
                this.Accent = System.BitConverter.ToInt32(this.Bytes, 20);
                this.AccentAndHyperlink = System.BitConverter.ToInt32(this.Bytes, 24);
                this.AccentAndFollowedHyperlink = System.BitConverter.ToInt32(this.Bytes, 28);
        }
    }
}
