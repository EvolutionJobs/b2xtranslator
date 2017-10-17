using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.DocFileFormat
{
    public class SectionDescriptor : ByteStructure
    {
        public short fn;
        public short fnMpr;
        public int fcMpr;

        /// <summary>
        /// A signed integer that specifies the position in the WordDocument Stream where a Sepx structure is located.
        /// </summary>
        public int fcSepx;

        private const int SED_LENGTH = 12;

        public SectionDescriptor(VirtualStreamReader reader, int length)
            : base(reader, length)
        {
            this.fn = this._reader.ReadInt16();
            this.fcSepx = this._reader.ReadInt32();
            this.fnMpr = this._reader.ReadInt16();
            this.fcMpr = this._reader.ReadInt32();
        }
    }
}
