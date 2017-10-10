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
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class TableInfo
    {
        /// <summary>
        /// 
        /// </summary>
        public bool fInTable;

        /// <summary>
        /// 
        /// </summary>
        public bool fTtp;

        /// <summary>
        /// 
        /// </summary>
        public bool fInnerTtp;

        /// <summary>
        /// 
        /// </summary>
        public bool fInnerTableCell;

        /// <summary>
        /// 
        /// </summary>
        public uint iTap;

        public TableInfo(ParagraphPropertyExceptions papx)
        {
            foreach (SinglePropertyModifier sprm in papx.grpprl)
            {
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPFInTable)
                {
                    this.fInTable = Utils.ByteToBool(sprm.Arguments[0]);
                }
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPFTtp)
                {
                    this.fTtp = Utils.ByteToBool(sprm.Arguments[0]);
                }
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPFInnerTableCell)
                {
                    this.fInnerTableCell = Utils.ByteToBool(sprm.Arguments[0]);
                }
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPFInnerTtp)
                {
                    this.fInnerTtp = Utils.ByteToBool(sprm.Arguments[0]);
                }
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPItap)
                {
                    this.iTap = System.BitConverter.ToUInt32(sprm.Arguments, 0);
                    if (this.iTap > 0)
                        this.fInTable = true;
                }
                if ((int)sprm.OpCode == 0x66A)
                {
                    //add value!
                    this.iTap = System.BitConverter.ToUInt32(sprm.Arguments, 0);
                    if (this.iTap > 0)
                        this.fInTable = true;
                }
            }
        }
    }
}
