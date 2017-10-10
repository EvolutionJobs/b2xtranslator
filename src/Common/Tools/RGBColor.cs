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

namespace DIaLOGIKa.b2xtranslator.Tools
{
    public class RGBColor
    {
        public enum ByteOrder 
        {
            RedFirst,
            RedLast
        }

        public byte Red;
        public byte Green;
        public byte Blue;
        public byte Alpha;
        public string SixDigitHexCode;
        public string EightDigitHexCode;

        public RGBColor(int cv, ByteOrder order)
        {
            var bytes = System.BitConverter.GetBytes(cv);

            if(order == ByteOrder.RedFirst)
            {
                //R
                this.Red = bytes[0];
                SixDigitHexCode = String.Format("{0:x2}", this.Red);
                //G
                this.Green = bytes[1];
                SixDigitHexCode += String.Format("{0:x2}", this.Green);
                //B
                this.Blue = bytes[2];
                SixDigitHexCode += String.Format("{0:x2}", this.Blue);
                EightDigitHexCode = SixDigitHexCode;
                //Alpha
                this.Alpha = bytes[3];
                EightDigitHexCode += String.Format("{0:x2}", this.Alpha);
            }
            else if (order == ByteOrder.RedLast)
            {
                //R
                this.Red = bytes[2];
                SixDigitHexCode = String.Format("{0:x2}", this.Red);
                //G
                this.Green = bytes[1];
                SixDigitHexCode += String.Format("{0:x2}", this.Green);
                //B
                this.Blue = bytes[0];
                SixDigitHexCode += String.Format("{0:x2}", this.Blue);
                EightDigitHexCode = SixDigitHexCode;
                //Alpha
                this.Alpha = bytes[3];
                EightDigitHexCode += String.Format("{0:x2}", this.Alpha);
            }

        }
    }
}
