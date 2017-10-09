using System.Diagnostics;
/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
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

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgErr : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgErr;

        public PtgErr(IStreamReader reader, PtgNumber ptgid)
            : base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 2;

            byte err = reader.ReadByte();
            this.Data = "";
            switch (err)
            {
                case 0x00:
                    this.Data = "#NULL!";
                    break;
                case 0x07:
                    this.Data = "#DIV/0!";
                    break;
                case 0x0F:
                    this.Data = "#VALUE!";
                    break;
                case 0x17:
                    this.Data = "#REF!";
                    break;
                case 0x1D:
                    this.Data = "#NAME?";
                    break;
                case 0x24:
                    this.Data = "#NUM!";
                    break;
                case 0x2A:
                    this.Data = "#N/A";
                    break;
            }
            this.type = PtgType.Operand;
            this.popSize = 1;
        }
    }
}
