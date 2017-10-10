/*
 * Copyright (c) 2009, DIaLOGIKa
 *
 * All rights reserved.
 * 
 * Redistribution and use in source and binary forms, with or without 
 * modification, are permitted provided that the following conditions are met:
 * 
 *     * Redistributions of source code must retain the above copyright 
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright 
 *       notice, this list of conditions and the following disclaimer in the 
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the names of copyright holders, nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software 
 *       without specific prior written permission.
 * 
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" 
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED 
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. 
 * IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, 
 * INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, 
 * BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, 
 * DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF 
 * LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE 
 * OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF 
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE. 
 */

using System;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies Boolean properties of the picture Obj containing this FtPioGrbit.
    /// </summary>
    public class FtPioGrbit
    {
        /// <summary>
        /// Reserved. MUST be 0x08.
        /// </summary>
        public ushort ft;

        /// <summary>
        /// Reserved. MUST be 0x02.
        /// </summary>
        public ushort cb;

        public bool fAutoPict;

        public bool fDde;

        public bool fPrintCalc;

        public bool fIcon;

        public bool fCtl;

        public bool fPrstm;

        public bool fCamera;

        public bool fDefaultSize;

        public bool fAutoLoad;

        public FtPioGrbit(IStreamReader reader)
        {
            this.ft = reader.ReadUInt16();
            this.cb = reader.ReadUInt16();

            var flags = reader.ReadUInt16();
            this.fAutoPict = Utils.BitmaskToBool(flags, 0x0001);
            this.fDde = Utils.BitmaskToBool(flags, 0x0002);
            this.fPrintCalc = Utils.BitmaskToBool(flags, 0x0004);
            this.fIcon = Utils.BitmaskToBool(flags, 0x0008);
            this.fCtl = Utils.BitmaskToBool(flags, 0x0010);
            this.fPrstm = Utils.BitmaskToBool(flags, 0x0020);

            this.fCamera = Utils.BitmaskToBool(flags, 0x0080);
            this.fDefaultSize = Utils.BitmaskToBool(flags, 0x0100);
            this.fAutoLoad = Utils.BitmaskToBool(flags, 0x0200);
        }
    }
}