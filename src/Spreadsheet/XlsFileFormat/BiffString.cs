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
using System;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    /// <summary>
    /// Implementation of a BIFF8 Unicode String
    /// 
    /// Excel 97 and later versions use unicode strings.  
    /// In BIFF8, strings are stored in a compressed format. 
    /// 
    /// Unicode strings usually require 2 bytes of storage per character.  
    /// Because most strings in USA/English Excel have all of the high bytes of unicode 
    /// characters set to 00h, the strings can be saved using a compressed unicode format.  
    /// </summary>
    public class BiffString
    {
        /// <summary>
        /// Count of characters in the string (Note: this is the number of characters, NOT the number of bytes)
        /// </summary>
        private ushort cch;

        /// <summary>
        /// Option flags
        /// </summary>
        private byte grbit;

        /// <summary>
        /// Array of string characters and formatting runs 
        /// </summary>
        private byte[] rgb;

        /// <summary>
        /// =0 if all the characters in the string have a high byte of 00h 
        ///    and only the low bytes are saved in the file (compressed)
        /// =1 if at least one character in the string has a nonzero high byte and 
        ///    therefore all characters in the string are saved as double-byte characters (not compressed)
        /// </summary>
        private bool fHighByte;

        /// <summary>
        /// Extended string follows (East Asian versions)
        /// </summary>
        private bool fExtSt;

        /// <summary>
        /// Rich string follows
        /// </summary>
        private bool fRichSt;
        
        public BiffString(IStreamReader reader)
        {
            cch = reader.ReadUInt16();
            grbit = reader.ReadByte();

            fHighByte = Utils.BitmaskToBool(grbit, 0x0001);
            fExtSt = Utils.BitmaskToBool(grbit, 0x0004);
            fRichSt = Utils.BitmaskToBool(grbit, 0x0008);
        }
    }
}
