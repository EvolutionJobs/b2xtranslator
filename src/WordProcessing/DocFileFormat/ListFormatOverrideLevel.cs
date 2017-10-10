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

using DIaLOGIKa.b2xtranslator.Tools;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class ListFormatOverrideLevel : ByteStructure
    {
        /// <summary>
        /// Start-at value if fFormatting==false and fStartAt==true. 
        /// If fFormatting == true, the start is stored in the LVL
        /// </summary>
        public int iStartAt;

        /// <summary>
        /// The level to be overridden
        /// </summary>
        public byte ilvl;

        /// <summary>
        /// True if the start-at value is overridden
        /// </summary>
        public bool fStartAt;

        /// <summary>
        /// True if the formatting is overridden
        /// </summary>
        public bool fFormatting;

        private const int LFOLVL_LENGTH = 6;

        /// <summary>
        /// Parses the bytes to retrieve a ListFormatOverrideLevel
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public ListFormatOverrideLevel(VirtualStreamReader reader, int length)
            : base(reader, length)
        {
            long startPos = _reader.BaseStream.Position;

            this.iStartAt = _reader.ReadInt32();
            int flag = (int)_reader.ReadInt16();
            this.ilvl = (byte)(flag & 0x000F);
            this.fStartAt = Utils.BitmaskToBool(flag, 0x0010);
            this.fFormatting = Utils.BitmaskToBool(flag, 0x0020);

            //it's a complete override, so the fix part is followed by LVL struct
            if (this.fFormatting)
            {

            }

            _reader.BaseStream.Seek(startPos, System.IO.SeekOrigin.Begin);
            _rawBytes = _reader.ReadBytes(LFOLVL_LENGTH);
        }
    }
}
