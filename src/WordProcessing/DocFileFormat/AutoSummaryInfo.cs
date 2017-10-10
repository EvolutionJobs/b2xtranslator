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
    public class AutoSummaryInfo
    {
        /// <summary>
        /// True if the ASUMYI is valid
        /// </summary>
        public bool fValid;

        /// <summary>
        /// True if AutoSummary View is active
        /// </summary>
        public bool fView;

        /// <summary>
        /// Display method for AutoSummary View: <br/>
        /// 0 = Emphasize in current doc<br/>
        /// 1 = Reduce doc to summary<br/>
        /// 2 = Insert into doc<br/>
        /// 3 = Show in new document
        /// </summary>
        public short iViewBy;

        /// <summary>
        /// True if File Properties summary information 
        /// should be updated after the next summarization
        /// </summary>
        public bool fUpdateProps;

        /// <summary>
        /// Dialog summary level
        /// </summary>
        public short wDlgLevel;

        /// <summary>
        /// Upper bound for lLevel for sentences in this document
        /// </summary>
        public int lHighestLevel;

        /// <summary>
        /// Show document sentences at or below this level
        /// </summary>
        public int lCurrentLevel;

        /// <summary>
        /// Parses the bytes to retrieve a AutoSummaryInfo
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public AutoSummaryInfo(byte[] bytes)
        {
            if (bytes.Length == 12)
            {
                //split byte 0 and 1 into bits
                var bits = new BitArray(new byte[] { bytes[0], bytes[1] });
                this.fValid = bits[0];
                this.fView = bits[1];
                this.iViewBy = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 2, 2));
                this.fUpdateProps = bits[4];

                this.wDlgLevel = System.BitConverter.ToInt16(bytes, 2);
                this.lHighestLevel = System.BitConverter.ToInt32(bytes, 4);
                this.lCurrentLevel = System.BitConverter.ToInt32(bytes, 8);
            }
            else
            {
                throw new ByteParseException("Cannot parse the struct ASUMYI, the length of the struct doesn't match");
            }
        }
    }
}
