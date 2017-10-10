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

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class DocumentTypographyInfo
    {
        /// <summary>
        /// True if we're kerning punctation
        /// </summary>
        public bool fKerningPunct;

        /// <summary>
        /// Kinsoku method of justification:<br/>
        /// 0 = always expand<br/>
        /// 1 = compress punctation<br/>
        /// 2 = compress punctation and kana
        /// </summary>
        public short iJustification;

        /// <summary>
        /// Level of kinsoku:<br/>
        /// 0 = level 1<br/>
        /// 1 = Level 2<br/>
        /// 2 = Custom
        /// </summary>
        public short iLevelOfKinsoku;

        /// <summary>
        /// "2 page on 1" feature is turned on
        /// </summary>
        public bool f2on1;

        /// <summary>
        /// Old East Asian feature
        /// </summary>
        public bool fOldDefineLineBaseOnGrid;

        /// <summary>
        /// Custom Kinsoku
        /// </summary>
        public short iCustomKsu;

        /// <summary>
        /// When set to true, use strict (level 2) Kinsoku rules
        /// </summary>
        public bool fJapaneseUseLevel2;

        /// <summary>
        /// Length of rgxchFPunct
        /// </summary>
        public short cchFollowingPunct;

        /// <summary>
        /// Length of rgxchLPunct
        /// </summary>
        public short cchLeadingPunct;

        /// <summary>
        /// Array of characters that should never appear at the start of a line
        /// </summary>
        public char[] rgxchFPunct;

        /// <summary>
        /// Array of characters that should never appear at the end of a line
        /// </summary>
        public char[] rgxchLPunct;

        /// <summary>
        /// Parses the bytes to retrieve a DocumentTypographyInfo
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public DocumentTypographyInfo(byte[] bytes)
        {
            if (bytes.Length == 310)
            {
                BitArray bits;

                //split byte 0 and 1 into bits
                bits = new BitArray(new byte[] { bytes[0], bytes[1] });
                this.fKerningPunct = bits[0];
                this.iJustification = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 1, 2));
                this.iLevelOfKinsoku = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 3, 2));
                this.f2on1 = bits[5];
                this.fOldDefineLineBaseOnGrid = bits[6];
                this.iCustomKsu = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 7, 3));
                this.fJapaneseUseLevel2 = bits[10];

                this.cchFollowingPunct = System.BitConverter.ToInt16(bytes, 2);
                this.cchLeadingPunct = System.BitConverter.ToInt16(bytes, 4);

                var fpunctBytes = new byte[202];
                Array.Copy(bytes, 6, fpunctBytes, 0, fpunctBytes.Length);
                this.rgxchFPunct = System.Text.Encoding.Unicode.GetString(fpunctBytes).ToCharArray();

                var lpunctBytes = new byte[102];
                Array.Copy(bytes, 208, lpunctBytes, 0, lpunctBytes.Length);
                this.rgxchLPunct = System.Text.Encoding.Unicode.GetString(lpunctBytes).ToCharArray();
            }
            else
            {
                throw new ByteParseException("Cannot parse the struct DOPTYPOGRAPHY, the length of the struct doesn't match");
            }
        }
    }
}
