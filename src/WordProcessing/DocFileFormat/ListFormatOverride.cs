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

using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class ListFormatOverride : ByteStructure
    {
        /// <summary>
        /// List ID of corresponding ListData
        /// </summary>
        public int lsid;

        /// <summary>
        /// Count of levels whose format is overridden
        /// </summary>
        public byte clfolvl;

        /// <summary>
        /// Specifies the field this LFO represents. 
        /// MUST be a value from the following table:<br/>
        /// 0x00:   This LFO is not used for any field.<br/>
        /// 0xFC:   This LFO is used for the AUTONUMLGL field.<br/>
        /// 0xFD:   This LFO is used for the AUTONUMOUT field.<br/>
        /// 0xFE:   This LFO is used for the AUTONUM field.<br/>
        /// 0xFF:   This LFO is not used for any field.
        /// </summary>
        public byte ibstFltAutoNum;

        /// <summary>
        /// A grfhic that specifies HTML incompatibilities.
        /// </summary>
        public byte grfhic;

        /// <summary>
        /// Array of all levels whose format is overridden
        /// </summary>
        public ListFormatOverrideLevel[] rgLfoLvl;

        private const int LFO_LENGTH = 16;

        /// <summary>
        /// Parses the given Stream Reader to retrieve a ListFormatOverride
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public ListFormatOverride(VirtualStreamReader reader, int length) : base(reader, length)
        {
            long startPos = _reader.BaseStream.Position;

            this.lsid = _reader.ReadInt32();
            _reader.ReadBytes(8);
            this.clfolvl = _reader.ReadByte();
            this.ibstFltAutoNum = _reader.ReadByte();
            this.grfhic = _reader.ReadByte();
            _reader.ReadByte();

            this.rgLfoLvl = new ListFormatOverrideLevel[this.clfolvl];

            _reader.BaseStream.Seek(startPos, System.IO.SeekOrigin.Begin);
            _rawBytes = _reader.ReadBytes(LFO_LENGTH);
        }
    }
}
