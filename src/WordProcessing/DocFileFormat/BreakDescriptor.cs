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
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class BreakDescriptor : ByteStructure
    {
        /// <summary>
        /// Except in textbox BKD, index to PGD in plfpgd that describes the page this break is on
        /// </summary>
        public short ipgd;

        /// <summary>
        /// In textbox BKD
        /// </summary>
        public short itxbxs;

        /// <summary>
        /// Number of cp's considered for this break; note that the CP's described by cpDepend in this break reside in the next BKD
        /// </summary>
        public short dcpDepend;

        /// <summary>
        /// 
        /// </summary>
        public ushort icol;

        /// <summary>
        /// When true, this indicates that this is a table break.
        /// </summary>
        public bool fTableBreak;

        /// <summary>
        /// When true, this indicates that this is a column break.
        /// </summary>
        public bool fColumnBreak;

        /// <summary>
        /// Used temporarily while Word is running.
        /// </summary>
        public bool fMarked;

        /// <summary>
        /// In textbox BKD, when true indicates cpLim of this textbox is not valid
        /// </summary>
        public bool fUnk;

        /// <summary>
        /// In textbox BKD, when true indicates that text overflows the end of this textbox
        /// </summary>
        public bool fTextOverflow;

        public BreakDescriptor(VirtualStreamReader reader, int length)
            : base(reader, length)
        {
            this.ipgd = reader.ReadInt16();
            this.itxbxs = this.ipgd;
            this.dcpDepend = reader.ReadInt16();
            int flag = (int)reader.ReadInt16();
            this.icol = (ushort)Utils.BitmaskToInt(flag, 0x00FF);
            this.fTableBreak = Utils.BitmaskToBool(flag, 0x0100);
            this.fColumnBreak = Utils.BitmaskToBool(flag, 0x0200);
            this.fMarked = Utils.BitmaskToBool(flag, 0x0400);
            this.fUnk = Utils.BitmaskToBool(flag, 0x0800);
            this.fTextOverflow = Utils.BitmaskToBool(flag, 0x1000);
        }
    }
}
