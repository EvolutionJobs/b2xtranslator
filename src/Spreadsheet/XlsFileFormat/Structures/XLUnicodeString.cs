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

using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies a Unicode string. 
    /// </summary>
    public class XLUnicodeString
    {
        /// <summary>
        /// An unsigned integer that specifies the count of characters in the string. 
        /// 
        /// MUST be equal to the number of characters in st.
        /// </summary>
        public ushort cch;

        /// <summary>
        /// A bit that specifies whether the characters in rgb are double-byte characters. 
        /// 
        /// MUST be a value from the following table: 
        ///     Value   Meaning
        ///     0x0     All the characters in the string have a high byte of 0x00 and only the low bytes are in rgb.
        ///     0x1     All the characters in the string are saved as double-byte characters in rgb.
        /// </summary>
        public bool fHighByte;

        /// <summary>
        /// An array of bytes that specifies the characters. 
        /// 
        /// If fHighByte is 0x0, the size of the array MUST be equal to the count of characters in the string. 
        /// If fHighByte is 0x1, the size of the array MUST be equal to 2 times the count of characters in the string.
        /// </summary>
        public byte[] rgb;        


        public XLUnicodeString(IStreamReader reader)
        {
            this.cch = reader.ReadUInt16();

            this.fHighByte = Utils.BitmaskToBool(reader.ReadByte(), 0x0001);

            if (fHighByte)
            {
                this.rgb = new byte[2 * cch];
            }
            else
            {
                this.rgb = new byte[cch];
            }

            this.rgb = reader.ReadBytes(this.rgb.Length);
        }

        public string Value
        {
            get
            {
                if (fHighByte)
                {
                    return Encoding.Unicode.GetString(this.rgb);
                }
                else
                {
                    return Encoding.GetEncoding(1252).GetString(this.rgb);
                }
            }
        }
    }
}
