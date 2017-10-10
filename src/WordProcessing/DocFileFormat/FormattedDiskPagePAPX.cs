using System;
using System.Collections.Generic;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.DocFileFormat
{
    public class FormattedDiskPagePAPX : FormattedDiskPage
    {
        /// <summary>
        /// An array of the BX data structure.<br/>
        /// BX is a 13 byte data structure. The first byte of each is the word offset of the PAPX.
        /// </summary>
        public BX[] rgbx;

        /// <summary>
        /// grppapx consists of all of the PAPXs stored in FKP concatenated end to end. 
        /// Each PAPX begins with a count of words which records its length padded to a word boundary.
        /// </summary>
        public ParagraphPropertyExceptions[] grppapx;

        public struct BX
        {
            public byte wordOffset;
            public ParagraphHeight phe;
        }

        public FormattedDiskPagePAPX(VirtualStream wordStream, int offset, VirtualStream dataStream)
        {
            this.Type = FKPType.Paragraph;
            this.WordStream = wordStream;

            //read the 512 bytes (FKP)
            var bytes = new byte[512];
            wordStream.Read(bytes, 0, 512, offset);

            //get the count
            this.crun = bytes[511];

            //create and fill the array with the adresses
            this.rgfc = new int[this.crun + 1];
            int j = 0;
            for (int i = 0; i < this.rgfc.Length; i++)
            {
                this.rgfc[i] = System.BitConverter.ToInt32(bytes, j);
                j += 4;
            }

            //create arrays
            this.rgbx = new BX[this.crun];
            this.grppapx = new ParagraphPropertyExceptions[this.crun];

            j = 4*(this.crun+1);
            for (int i = 0; i < this.rgbx.Length; i++)
            {
                //read the 12 for PHE
                var phe = new byte[12];
                Array.Copy(bytes, j+1, phe, 0, phe.Length);

                //fill the rgbx array
                var bx = new BX
                {
                    wordOffset = bytes[j],
                    phe = new ParagraphHeight(phe, false)
                };
                this.rgbx[i] = bx;
                j += 13;

                if (bx.wordOffset != 0)
                {
                    //read first byte of PAPX
                    //PAPX is stored in a FKP; so the first byte is a count of words
                    byte padbyte = 0;
                    byte cw = bytes[bx.wordOffset * 2];

                    //if that byte is zero, it's a pad byte, and the word count is the following byte
                    if (cw == 0)
                    {
                        padbyte = 1;
                        cw = bytes[bx.wordOffset * 2 + 1];
                    }

                    if (cw != 0)
                    {
                        //read the bytes for papx
                        var papx = new byte[cw * 2];
                        Array.Copy(bytes, (bx.wordOffset * 2) + padbyte + 1, papx, 0, papx.Length);

                        //parse PAPX and fill grppapx
                        this.grppapx[i] = new ParagraphPropertyExceptions(papx, dataStream);
                    }
                }
                else
                {
                    //create a PAPX which doesn't modify anything
                    this.grppapx[i] = new ParagraphPropertyExceptions();
                }
            }
        }

        /// <summary>
        /// Parses the 0Table (or 1Table) for FKP _entries containing PAPX
        /// </summary>
        /// <param name="fib">The FileInformationBlock</param>
        /// <param name="wordStream">The WordDocument stream</param>
        /// <param name="tableStream">The 0Table stream</param>
        /// <returns></returns>
        public static List<FormattedDiskPagePAPX> GetAllPAPXFKPs(FileInformationBlock fib, VirtualStream wordStream, VirtualStream tableStream, VirtualStream dataStream)
        {
            var list = new List<FormattedDiskPagePAPX>();

            //get bintable for PAPX
            var binTablePapx = new byte[fib.lcbPlcfBtePapx];
            tableStream.Read(binTablePapx, 0, binTablePapx.Length, (int)fib.fcPlcfBtePapx);

            //there are n offsets and n-1 fkp's in the bin table
            int n = (((int)fib.lcbPlcfBtePapx - 4) / 8) + 1;

            //Get the indexed PAPX FKPs
            for (int i = (n * 4); i < binTablePapx.Length; i += 4)
            {
                //indexed FKP is the xth 512byte page
                int fkpnr = System.BitConverter.ToInt32(binTablePapx, i);

                //so starts at:
                int offset = fkpnr * 512;

                //parse the FKP and add it to the list
                list.Add(new FormattedDiskPagePAPX(wordStream, offset, dataStream));
            }

            return list;
        }

        /// <summary>
        /// Returns a list of all PAPX FCs between they given boundaries.
        /// </summary>
        /// <param name="fcMin">The lower boundary</param>
        /// <param name="fcMax">The upper boundary</param>
        /// <param name="fib">The FileInformationBlock</param>
        /// <param name="wordStream">The VirtualStream "WordStream"</param>
        /// <param name="tableStream">The VirtualStream "0Table" or "1Table"</param>
        /// <returns>The FCs</returns>
        public static List<int> GetFileCharacterPositions(
            int fcMin,
            int fcMax,
            FileInformationBlock fib, 
            VirtualStream wordStream, 
            VirtualStream tableStream,
            VirtualStream dataStream)
        {
            var list = new List<int>();
            var fkps = FormattedDiskPagePAPX.GetAllPAPXFKPs(fib, wordStream, tableStream, dataStream);

            for (int i = 0; i < fkps.Count; i++ )
            {
                FormattedDiskPage fkp = fkps[i];

                //the last entry of each is always the same as the first entry of the next FKP
                //so, ignore all last _entries except for the last FKP.
                int max = fkp.rgfc.Length;
                if (i < fkps.Count - 1)
                    max--;

                for (int j = 0; j < max; j++)
                {
                    if(fkp.rgfc[j] >= fcMin && fkp.rgfc[j] < fcMax)
                        list.Add(fkp.rgfc[j]);
                }
            }

            return list;
        }

        /// <summary>
        /// Returnes a list of all ParagraphPropertyExceptions which correspond to text 
        /// between the given offsets.
        /// </summary>
        /// <param name="fcMin">The lower boundary</param>
        /// <param name="fcMax">The upper boundary</param>
        /// <param name="fib">The FileInformationBlock</param>
        /// <param name="wordStream">The VirtualStream "WordStream"</param>
        /// <param name="tableStream">The VirtualStream "0Table" or "1Table"</param>
        /// <returns>The FCs</returns>
        public static List<ParagraphPropertyExceptions> GetParagraphPropertyExceptions(
            int fcMin,
            int fcMax,
            FileInformationBlock fib, 
            VirtualStream wordStream, 
            VirtualStream tableStream,
            VirtualStream dataStream)
        {
            var list = new List<ParagraphPropertyExceptions>();
            var fkps = FormattedDiskPagePAPX.GetAllPAPXFKPs(fib, wordStream, tableStream, dataStream);

            for (int i = 0; i < fkps.Count; i++)
            {
                var fkp = fkps[i];

                for (int j = 0; j < fkp.grppapx.Length; j++)
                {
                    if (fkp.rgfc[j] >= fcMin && fkp.rgfc[j] < fcMax)
                        list.Add(fkp.grppapx[j]);
                }
            }

            return list;
        }

    }
}
