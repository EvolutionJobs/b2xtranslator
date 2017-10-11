using System.Collections.Generic;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.DocFileFormat
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

            this.FirstHeaders = new List<CharacterRange>();
            this.EvenHeaders = new List<CharacterRange>();
            this.OddHeaders = new List<CharacterRange>();
            this.FirstFooters = new List<CharacterRange>();
            this.EvenFooters = new List<CharacterRange>();
            this.OddFooters = new List<CharacterRange>();

            //read the Table
            var table = new int[doc.FIB.lcbPlcfHdd / 4];
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
