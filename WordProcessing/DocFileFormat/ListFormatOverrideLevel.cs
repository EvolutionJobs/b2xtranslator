using b2xtranslator.Tools;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.DocFileFormat
{
    public class ListFormatOverrideLevel : ByteStructure
    {
        /// <summary>
        /// Start-at value if fFormatting==false and fStartAt==true. 
        /// If fFormatting == true, the start is stored in the LVL
        /// </summary>
        public int iStartAt;

        /// <summary>
        /// The level to be overridden
        /// </summary>
        public byte ilvl;

        /// <summary>
        /// True if the start-at value is overridden
        /// </summary>
        public bool fStartAt;

        /// <summary>
        /// True if the formatting is overridden
        /// </summary>
        public bool fFormatting;

        private const int LFOLVL_LENGTH = 6;

        /// <summary>
        /// Parses the bytes to retrieve a ListFormatOverrideLevel
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public ListFormatOverrideLevel(VirtualStreamReader reader, int length)
            : base(reader, length)
        {
            long startPos = this._reader.BaseStream.Position;

            this.iStartAt = this._reader.ReadInt32();
            int flag = (int)this._reader.ReadInt16();
            this.ilvl = (byte)(flag & 0x000F);
            this.fStartAt = Utils.BitmaskToBool(flag, 0x0010);
            this.fFormatting = Utils.BitmaskToBool(flag, 0x0020);

            //it's a complete override, so the fix part is followed by LVL struct
            if (this.fFormatting)
            {

            }

            this._reader.BaseStream.Seek(startPos, System.IO.SeekOrigin.Begin);
            this._rawBytes = this._reader.ReadBytes(LFOLVL_LENGTH);
        }
    }
}
