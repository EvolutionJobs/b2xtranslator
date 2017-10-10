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
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class AutoNumberedListDataDescriptor
    {
        public AutoNumberLevelDescriptor anlv;

        /// <summary>
        /// Number only 1 item per table cell
        /// </summary>
        public bool fNumber1;

        /// <summary>
        /// Number across cells in table rows
        /// </summary>
        public bool fNumberAcross;

        /// <summary>
        /// Restart heading number on section boundary
        /// </summary>
        public bool fRestartHdn;

        /// <summary>
        /// Not used
        /// </summary>
        public bool fSpareX;
	
        /// <summary>
        /// Characters displayed before/after auto number
        /// </summary>
        public char[] rgxch;

        /// <summary>
        /// Creates a new AutoNumberedListDataDescriptor with defaut values
        /// </summary>
        public AutoNumberedListDataDescriptor()
        {
            setDefaultValues();
        }

        /// <summary>
        /// Parses the bytes to retrieve a AutoNumberedListDataDescriptor
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public AutoNumberedListDataDescriptor(byte[] bytes)
        {
            if (bytes.Length == 88)
            {
                //copies the first 20 bytes into a new array
                var anlvArray = new byte[20];
                Array.Copy(bytes, anlvArray, anlvArray.Length);
                //parse these bytes 
                this.anlv = new AutoNumberLevelDescriptor(anlvArray);

                //parse the rest
                if (bytes[20] == 1)
                    this.fNumber1 = true;
                if (bytes[21] == 1)
                    this.fNumberAcross = true;
                if (bytes[22] == 1)
                    this.fRestartHdn = true;
                if (bytes[23] == 1)
                    this.fSpareX = true;

                this.rgxch = new char[32];
                int j=0;
                for (int i = 24; i <= 88; i += 2)
                {
                    rgxch[j] = Convert.ToChar(BitConverter.ToInt16(bytes, i));
                    j++;
                }
            }
            else
            {
                throw new ByteParseException(
                    "Cannot parse the struct ANLD, the length of the struct doesn't match");
            }
        }

        private void setDefaultValues()
        {
            this.fNumber1 = false;
            this.fNumberAcross = false;
            this.fRestartHdn = false;
            this.fSpareX = false;
            this.rgxch = Utils.ClearCharArray(new char[32]);
        }
    }
}
