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
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class FormattedDiskPageCHPX : FormattedDiskPage
    {
        /// <summary>
        /// An array of bytes where each byte is the word offset of a CHPX.
        /// </summary>
        public byte[] rgb;

        /// <summary>
        /// Consists all of the CHPXs stored in this FKP.
        /// </summary>
        public CharacterPropertyExceptions[] grpchpx;

        public FormattedDiskPageCHPX(VirtualStream wordStream, int offset)
        {
            this.Type = FKPType.Character;
            this.WordStream = wordStream;

            //read the 512 bytes (FKP)
            var bytes = new byte[512];
            wordStream.Read(bytes, 0, 512, offset);

            //get the count first
            this.crun = bytes[511];

            //create and fill the array with the adresses
            this.rgfc = new int[this.crun + 1];
            int j = 0;
            for (int i = 0; i < rgfc.Length; i++)
            {
                rgfc[i] = System.BitConverter.ToInt32(bytes, j);
                j += 4;
            }

            //create arrays
            this.rgb = new byte[this.crun];
            this.grpchpx = new CharacterPropertyExceptions[this.crun];

            j = 4*(this.crun+1);
            for (int i = 0; i < rgb.Length; i++)
            {
                //fill the rgb array
                byte wordOffset = bytes[j];
                this.rgb[i] = wordOffset;
                j++;

                if (wordOffset != 0)
                {
                    //read first byte of CHPX
                    //it's the count of bytes
                    byte cb = bytes[wordOffset * 2];

                    //read the bytes of chpx
                    var chpx = new byte[cb];
                    Array.Copy(bytes, (wordOffset * 2) + 1, chpx, 0, chpx.Length);

                    //parse CHPX and fill grpchpx
                    grpchpx[i] = new CharacterPropertyExceptions(chpx);
                }
                else
                {
                    //create a CHPX which doesn't modify anything
                    grpchpx[i] = new CharacterPropertyExceptions();
                }
            }

        }

        /// <summary>
        /// Parses the 0Table (or 1Table) for FKP _entries containing CHPX
        /// </summary>
        /// <param name="fib">The FileInformationBlock</param>
        /// <param name="wordStream">The WordDocument stream</param>
        /// <param name="tableStream">The 0Table stream</param>
        /// <returns></returns>
        public static List<FormattedDiskPageCHPX> GetAllCHPXFKPs(FileInformationBlock fib, VirtualStream wordStream, VirtualStream tableStream)
        {
            var list = new List<FormattedDiskPageCHPX>();

            //get bintable for CHPX
            var binTableChpx = new byte[fib.lcbPlcfBteChpx];
            tableStream.Read(binTableChpx, 0, binTableChpx.Length, (int)fib.fcPlcfBteChpx);

            //there are n offsets and n-1 fkp's in the bin table
            int n = (((int)fib.lcbPlcfBteChpx - 4) / 8) + 1;

            //Get the indexed CHPX FKPs
            for (int i = (n * 4); i < binTableChpx.Length; i += 4)
            {
                //indexed FKP is the 6th 512byte page
                int fkpnr = System.BitConverter.ToInt32(binTableChpx, i);

                //so starts at:
                int offset = fkpnr * 512;

                //parse the FKP and add it to the list
                list.Add(new FormattedDiskPageCHPX(wordStream, offset));
            }

            return list;
        }
    }
}
