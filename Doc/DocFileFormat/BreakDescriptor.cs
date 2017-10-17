using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.DocFileFormat
{
    public class BreakDescriptor : ByteStructure
    {
        /// <summary>
        /// Except in textbox BKD, index to PGD in plfpgd that describes the page this break is on
        /// </summary>
        public short ipgd;

        /// <summary>
        /// In textbox BKD
        /// </summary>
        public short itxbxs;

        /// <summary>
        /// Number of cp's considered for this break; note that the CP's described by cpDepend in this break reside in the next BKD
        /// </summary>
        public short dcpDepend;

        /// <summary>
        /// 
        /// </summary>
        public ushort icol;

        /// <summary>
        /// When true, this indicates that this is a table break.
        /// </summary>
        public bool fTableBreak;

        /// <summary>
        /// When true, this indicates that this is a column break.
        /// </summary>
        public bool fColumnBreak;

        /// <summary>
        /// Used temporarily while Word is running.
        /// </summary>
        public bool fMarked;

        /// <summary>
        /// In textbox BKD, when true indicates cpLim of this textbox is not valid
        /// </summary>
        public bool fUnk;

        /// <summary>
        /// In textbox BKD, when true indicates that text overflows the end of this textbox
        /// </summary>
        public bool fTextOverflow;

        public BreakDescriptor(VirtualStreamReader reader, int length)
            : base(reader, length)
        {
            this.ipgd = reader.ReadInt16();
            this.itxbxs = this.ipgd;
            this.dcpDepend = reader.ReadInt16();
            int flag = (int)reader.ReadInt16();
            this.icol = (ushort)Utils.BitmaskToInt(flag, 0x00FF);
            this.fTableBreak = Utils.BitmaskToBool(flag, 0x0100);
            this.fColumnBreak = Utils.BitmaskToBool(flag, 0x0200);
            this.fMarked = Utils.BitmaskToBool(flag, 0x0400);
            this.fUnk = Utils.BitmaskToBool(flag, 0x0800);
            this.fTextOverflow = Utils.BitmaskToBool(flag, 0x1000);
        }
    }
}
