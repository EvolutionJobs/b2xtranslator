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
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class OutlineLiSTData
    {
        /// <summary>
        /// An array of  ANLV structures describing how heading numbers 
        /// should be displayed fpr each of Word's 0 outline heading levels
        /// </summary>
        public AutoNumberLevelDescriptor[] rganlv;

        /// <summary>
        /// When true, restart heading on section break
        /// </summary>
        public bool fRestartHdr;

        /// <summary>
        /// Reserved
        /// </summary>
        public bool fSpareOlst2;

        /// <summary>
        /// Reserved
        /// </summary>
        public bool fSpareOlst3;

        /// <summary>
        /// Reserved
        /// </summary>
        public bool fSpareOlst4;

        /// <summary>
        /// Text before/after number
        /// </summary>
        public char[] rgxch;

        /// <summary>
        /// Creates a new OutlineLiSTData with default values
        /// </summary>
        public OutlineLiSTData()
        {
            setDefaultValues();
        }

        /// <summary>
        /// Parses the bytes to retrieve a OutlineLiSTData
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public OutlineLiSTData(byte[] bytes)
        {
            if (bytes.Length == 248)
            {
                //Fill the rganlv array
                int j = 0;
                for (int i = 0; i < 180; i += 20)
                {
                    //copy the 20 byte pages
                    var page = new byte[20];
                    Array.Copy(bytes, i, page, 0, 20);
                    this.rganlv[j] = new AutoNumberLevelDescriptor(page);
                    j++;
                }

                //Set the flags
                this.fRestartHdr = Utils.IntToBool((int)bytes[180]);
                this.fSpareOlst2 = Utils.IntToBool((int)bytes[181]);
                this.fSpareOlst3 = Utils.IntToBool((int)bytes[182]);
                this.fSpareOlst4 = Utils.IntToBool((int)bytes[183]);

                //Fill the rgxch array
                j = 0;
                for (int i = 184; i < 64; i += 2)
                {
                    this.rgxch[j] = Convert.ToChar(System.BitConverter.ToInt16(bytes, i));
                    j++;
                }
            }
            else
            {
                throw new ByteParseException("Cannot parse the struct OLST, the length of the struct doesn't match");
            }
        }

        private void setDefaultValues()
        {
            this.fRestartHdr = false;
            this.fSpareOlst2 = false;
            this.fSpareOlst3 = false;
            this.fSpareOlst4 = false;
            this.rganlv = new AutoNumberLevelDescriptor[9];
            for (int i = 0; i < 9; i++)
            {
                this.rganlv[i] = new AutoNumberLevelDescriptor();
            }
            this.rgxch = Utils.ClearCharArray(new char[32]);
        }
    }
}
