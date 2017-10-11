using System.Collections.Generic;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.DocFileFormat
{
    public class ListFormatOverrideTable : List<ListFormatOverride>
    {
        private const int LFO_LENGTH = 16;
        private const int LFOLVL_LENGTH = 6;

        public ListFormatOverrideTable(FileInformationBlock fib, VirtualStream tableStream)
        {
            if (fib.lcbPlfLfo > 0)
            {
                var reader = new VirtualStreamReader(tableStream);
                reader.BaseStream.Seek((long)fib.fcPlfLfo, System.IO.SeekOrigin.Begin);

                //read the count of LFOs
                int count = reader.ReadInt32();

                //read the LFOs
                for (int i = 0; i < count; i++)
                {
                    this.Add(new ListFormatOverride(reader, LFO_LENGTH));
                }

                //read the LFOLVLs
                for (int i = 0; i < count; i++)
                {
                    for (int j = 0; j < this[i].clfolvl; j++)
                    {
                        this[i].rgLfoLvl[j] = new ListFormatOverrideLevel(reader, LFOLVL_LENGTH);
                    }
                }
            }
        }
    }
}
