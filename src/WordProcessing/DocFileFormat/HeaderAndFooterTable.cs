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
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class HeaderAndFooterTable
    {
        public List<CharacterRange> FirstHeaders;
        public List<CharacterRange> EvenHeaders;
        public List<CharacterRange> OddHeaders;
        public List<CharacterRange> FirstFooters;
        public List<CharacterRange> EvenFooters;
        public List<CharacterRange> OddFooters;

        public HeaderAndFooterTable(WordDocument doc)
        {
            IStreamReader tableReader = new VirtualStreamReader(doc.TableStream);

            FirstHeaders = new List<CharacterRange>();
            EvenHeaders = new List<CharacterRange>();
            OddHeaders = new List<CharacterRange>();
            FirstFooters = new List<CharacterRange>();
            EvenFooters = new List<CharacterRange>();
            OddFooters = new List<CharacterRange>();

            //read the Table
            Int32[] table = new Int32[doc.FIB.lcbPlcfHdd / 4];
            doc.TableStream.Seek(doc.FIB.fcPlcfHdd, System.IO.SeekOrigin.Begin);
            for (int i = 0; i < table.Length; i++)
            {
                table[i] = tableReader.ReadInt32();
            }
            int count = (table.Length - 8) / 6;

            int initialPos = doc.FIB.ccpText + doc.FIB.ccpFtn;

            //the first 6 _entries are about footnote and endnote formatting
            //so skip these _entries
            int pos = 6;
            for (int i = 0; i < count; i++)
            {
                //Even Header
                if (table[pos] == table[pos + 1])
                {
                    this.EvenHeaders.Add(null);
                }
                else
                {
                    this.EvenHeaders.Add(new CharacterRange(initialPos + table[pos], table[pos + 1] - table[pos]));
                }
                pos++;

                //Odd Header
                if (table[pos] == table[pos + 1])
                {
                    this.OddHeaders.Add(null);
                }
                else
                {
                    this.OddHeaders.Add(new CharacterRange(initialPos + table[pos], table[pos + 1] - table[pos]));
                }
                pos++;

                //Even Footer
                if (table[pos] == table[pos + 1])
                {
                    this.EvenFooters.Add(null);
                }
                else
                {
                    this.EvenFooters.Add(new CharacterRange(initialPos + table[pos], table[pos + 1] - table[pos]));
                }
                pos++;

                //Odd Footer
                if (table[pos] == table[pos + 1])
                {
                    this.OddFooters.Add(null);
                }
                else
                {
                    this.OddFooters.Add(new CharacterRange(initialPos + table[pos], table[pos + 1] - table[pos]));
                }
                pos++;

                //First Page Header
                if (table[pos] == table[pos + 1])
                {
                    this.FirstHeaders.Add(null);
                }
                else
                {
                    this.FirstHeaders.Add(new CharacterRange(initialPos + table[pos], table[pos + 1] - table[pos]));
                }
                pos++;

                //First Page Footers
                if (table[pos] == table[pos + 1])
                {
                    this.FirstFooters.Add(null);
                }
                else
                {
                    this.FirstFooters.Add(new CharacterRange(initialPos + table[pos], table[pos + 1] - table[pos]));
                }
                pos++;
            }
        }
    }
}
