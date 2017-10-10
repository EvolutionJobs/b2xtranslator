using b2xtranslator.Tools;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.DocFileFormat
{
    public class ListData : ByteStructure
    {
        /// <summary>
        /// Unique List ID
        /// </summary>
        public int lsid;

        /// <summary>
        /// Unique template code
        /// </summary>
        public int tplc;

        /// <summary>
        /// Array of shorts containing the istd‘s linked to each level of the list, 
        /// or ISTD_NIL (4095) if no style is linked.
        /// </summary>
        public short[] rgistd;

        /// <summary>
        /// True if this is a simple (one-level) list.<br/>
        /// False if this is a multilevel (nine-level) list.
        /// </summary>
        public bool fSimpleList;

        /// <summary>
        /// Word 6.0 compatibility option:<br/>
        /// True if the list should start numbering over at the beginning of each section.
        /// </summary>
        public bool fRestartHdn;

        /// <summary>
        /// To emulate Word 6.0 numbering: <br/>
        /// True if Auto numbering
        /// </summary>
        public bool fAutoNum;

        /// <summary>
        /// When true, this list was there before we started reading RTF.
        /// </summary>
        public bool fPreRTF;

        /// <summary>
        /// When true, list is a hybrid multilevel/simple (UI=simple, internal=multilevel)
        /// </summary>
        public bool fHybrid;

        /// <summary>
        /// Array of ListLevel describing the several levels of the list.
        /// </summary>
        public ListLevel[] rglvl;

        /// <summary>
        /// A grfhic that specifies HTML incompatibilities of the list.
        /// </summary>
        public byte grfhic;

        public const short ISTD_NIL = 4095;
        private const int LSTF_LENGTH = 28;


        /// <summary>
        /// Parses the StreamReader to retrieve a ListData
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public ListData(VirtualStreamReader reader, int length) : base(reader, length)
        {
            long startPos = this._reader.BaseStream.Position;

            this.lsid = this._reader.ReadInt32();
            this.tplc = this._reader.ReadInt32();

            this.rgistd = new short[9];
            for (int i = 0; i < 9; i++)
            {
                this.rgistd[i] = this._reader.ReadInt16();
            }

            //parse flagbyte
            int flag = (int)this._reader.ReadByte();
            this.fSimpleList = Utils.BitmaskToBool(flag, 0x01);

            if (this.fSimpleList)
                this.rglvl = new ListLevel[1];
            else
                this.rglvl = new ListLevel[9];

            this.fRestartHdn = Utils.BitmaskToBool(flag, 0x02);
            this.fAutoNum = Utils.BitmaskToBool(flag, 0x04);
            this.fPreRTF = Utils.BitmaskToBool(flag, 0x08);
            this.fHybrid = Utils.BitmaskToBool(flag, 0x10);

            this.grfhic = this._reader.ReadByte();

            this._reader.BaseStream.Seek(startPos, System.IO.SeekOrigin.Begin);
            this._rawBytes = this._reader.ReadBytes(LSTF_LENGTH);
        }
    }
}
