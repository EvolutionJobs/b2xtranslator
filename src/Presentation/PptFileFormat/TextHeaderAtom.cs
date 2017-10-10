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
using System.IO;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    public enum TextType
    {
        Title = 0,
        Body,
        Notes,
        Outline,
        Other,
        CenterBody,
        CenterTitle,
        HalfBody,
        QuarterBody
    };

    [OfficeRecord(3999)]
    public class TextHeaderAtom : Record
    {
        public TextType TextType;

        public TextAtom TextAtom;

        public TextStyleAtom TextStyleAtom;

        public TextHeaderAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.TextType = (TextType) this.Reader.ReadUInt32();
        }

        public void HandleTextDataRecord(ITextDataRecord tdRecord)
        {
            tdRecord.TextHeaderAtom = this;

            var textAtom = tdRecord as TextAtom;
            var tsAtom = tdRecord as TextStyleAtom;

            if (textAtom != null)
            {
                this.TextAtom = textAtom;
            }
            else if (tsAtom != null)
            {
                this.TextStyleAtom = tsAtom;
            }
            else
            {
                throw new NotImplementedException("Unhandled text data record type " + tdRecord.GetType().ToString());
            }
        }
    }

}
