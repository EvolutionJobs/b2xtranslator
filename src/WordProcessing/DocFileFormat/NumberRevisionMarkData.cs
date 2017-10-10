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
    public class NumberRevisionMarkData
    {
        /// <summary>
        /// True if this paragraph was numbered when revision 
        /// mark tracking was turned on
        /// </summary>
        public bool fNumRM;

        /// <summary>
        /// Index of author IDs stored in hsttbfRMark for the 
        /// paragraph number change
        /// </summary>
        public short ibstNumRM;

        /// <summary>
        /// Date of the paragraph number change
        /// </summary>
        public DateAndTime dttmNumRM;

        /// <summary>
        /// Index into xst of the locations of paragraph number 
        /// place holders for each level
        /// </summary>
        public char[] rgbxchNums;

        /// <summary>
        /// Number format code for the paragraph number 
        /// place holders for each level
        /// </summary>
        public char[] rgnfc;

        /// <summary>
        /// Numeric value for each place holder in xst
        /// </summary>
        public int[] PNBR;

        /// <summary>
        /// The text string for the paragraph number, 
        /// containing level place holders
        /// </summary>
        public char[] xst;

        /// <summary>
        /// Creates a new NumberRevisionMarkData with default values
        /// </summary>
        public NumberRevisionMarkData()
        {
            setDefaultValues();
        }

        /// <summary>
        /// Parses the bytes to retrieve a NumberRevisionMarkData
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public NumberRevisionMarkData(byte[] bytes)
        {
            if (bytes.Length == 128)
            {
                this.fNumRM = Utils.IntToBool((int)bytes[0]);
                this.ibstNumRM = System.BitConverter.ToInt16(bytes, 2);

                //copy date to new array and parse it
                var dttm = new byte[4];
                Array.Copy(bytes,4, dttm,0, 4);
                this.dttmNumRM = new DateAndTime(dttm);

                //fill the rgbxchNums char array
                this.rgbxchNums = new char[9];
                int j = 0;
                for (int i = 8; i < 17; i++)
                {
                    this.rgbxchNums[j] = Convert.ToChar(bytes[i]);
                    j++;
                }

                //fill the rgnfc char array
                this.rgnfc = new char[9];
                j = 0;
                for (int i = 17; i < 26; i++)
                {
                    this.rgnfc[j] = Convert.ToChar(bytes[i]);
                    j++;
                }

                //fill the PNBR array
                this.PNBR = new int[9];
                j = 0;
                for (int i = 28; i < 64; i+=4)
                {
                    this.PNBR[j] = System.BitConverter.ToInt32(bytes, i);
                    j++;
                }

                //fill the xst char array
                this.xst = new char[32];
                j = 0;
                for (int i = 64; i < 128; i += 2)
                {
                    this.xst[j] = Convert.ToChar(System.BitConverter.ToUInt16(bytes, i));
                    j++;
                }
            }
            else
            {
                throw new ByteParseException("Cannot parse the struct NUMRM, the length of the struct doesn't match");
            }
        }

        private void setDefaultValues()
        {
            this.dttmNumRM = new DateAndTime();
            this.fNumRM = false;
            this.ibstNumRM = 0;
            this.PNBR = Utils.ClearIntArray(new int[9]);
            this.rgbxchNums = Utils.ClearCharArray(new char[9]);
            this.rgnfc = Utils.ClearCharArray(new char[9]);
            this.xst = Utils.ClearCharArray(new char[32]);
        }
    }
}
