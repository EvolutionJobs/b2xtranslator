using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.DocFileFormat
{
    public class ListFormatOverride : ByteStructure
    {
        /// <summary>
        /// List ID of corresponding ListData
        /// </summary>
        public int lsid;

        /// <summary>
        /// Count of levels whose format is overridden
        /// </summary>
        public byte clfolvl;

        /// <summary>
        /// Specifies the field this LFO represents. 
        /// MUST be a value from the following table:<br/>
        /// 0x00:   This LFO is not used for any field.<br/>
        /// 0xFC:   This LFO is used for the AUTONUMLGL field.<br/>
        /// 0xFD:   This LFO is used for the AUTONUMOUT field.<br/>
        /// 0xFE:   This LFO is used for the AUTONUM field.<br/>
        /// 0xFF:   This LFO is not used for any field.
        /// </summary>
        public byte ibstFltAutoNum;

        /// <summary>
        /// A grfhic that specifies HTML incompatibilities.
        /// </summary>
        public byte grfhic;

        /// <summary>
        /// Array of all levels whose format is overridden
        /// </summary>
        public ListFormatOverrideLevel[] rgLfoLvl;

        private const int LFO_LENGTH = 16;

        /// <summary>
        /// Parses the given Stream Reader to retrieve a ListFormatOverride
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public ListFormatOverride(VirtualStreamReader reader, int length) : base(reader, length)
        {
            long startPos = this._reader.BaseStream.Position;

            this.lsid = this._reader.ReadInt32();
            this._reader.ReadBytes(8);
            this.clfolvl = this._reader.ReadByte();
            this.ibstFltAutoNum = this._reader.ReadByte();
            this.grfhic = this._reader.ReadByte();
            this._reader.ReadByte();

            this.rgLfoLvl = new ListFormatOverrideLevel[this.clfolvl];

            this._reader.BaseStream.Seek(startPos, System.IO.SeekOrigin.Begin);
            this._rawBytes = this._reader.ReadBytes(LFO_LENGTH);
        }
    }
}
