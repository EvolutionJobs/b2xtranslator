using b2xtranslator.OfficeDrawing;
using System.IO;
using b2xtranslator.Tools;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(4011)]
    public class DefaultRulerAtom : Record
    {
        public uint flags;

        public bool fDefaultTabSize;
        public bool fCLevels;
        public bool fTabStops;
        public bool fLeftMargin1;
        public bool fLeftMargin2;
        public bool fLeftMargin3;
        public bool fLeftMargin4;
        public bool fLeftMargin5;
        public bool fIndent1;
        public bool fIndent2;
        public bool fIndent3;
        public bool fIndent4;
        public bool fIndent5;

        public int cLevels;
        public int defaultTabSize;
        public TabStops tabs;
        public int leftMargin1;
        public int leftMargin2;
        public int leftMargin3;
        public int leftMargin4;
        public int leftMargin5;
        public int indent1;
        public int indent2;
        public int indent3;
        public int indent4;
        public int indent5;

        public DefaultRulerAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.flags = this.Reader.ReadUInt32();
            this.fDefaultTabSize = Utils.BitmaskToBool(this.flags, 0x1);
            this.fCLevels = Utils.BitmaskToBool(this.flags, 0x1 << 1);
            this.fTabStops = Utils.BitmaskToBool(this.flags, 0x1 << 2);
            this.fLeftMargin1 = Utils.BitmaskToBool(this.flags, 0x1 << 3);
            this.fLeftMargin2 = Utils.BitmaskToBool(this.flags, 0x1 << 4);
            this.fLeftMargin3 = Utils.BitmaskToBool(this.flags, 0x1 << 5);
            this.fLeftMargin4 = Utils.BitmaskToBool(this.flags, 0x1 << 6);
            this.fLeftMargin5 = Utils.BitmaskToBool(this.flags, 0x1 << 7);
            this.fIndent1 = Utils.BitmaskToBool(this.flags, 0x1 << 8);
            this.fIndent2 = Utils.BitmaskToBool(this.flags, 0x1 << 9);
            this.fIndent3 = Utils.BitmaskToBool(this.flags, 0x1 << 10);
            this.fIndent4 = Utils.BitmaskToBool(this.flags, 0x1 << 11);
            this.fIndent5 = Utils.BitmaskToBool(this.flags, 0x1 << 12);

            if (this.fCLevels) this.cLevels = this.Reader.ReadInt16();
            if (this.fDefaultTabSize) this.defaultTabSize = this.Reader.ReadInt16();

            if (this.fTabStops) this.tabs = new TabStops(this.Reader);

            if (this.fLeftMargin1) this.leftMargin1 = this.Reader.ReadInt16();
            if (this.fIndent1) this.indent1 = this.Reader.ReadInt16();
            if (this.fLeftMargin2) this.leftMargin2 = this.Reader.ReadInt16();
            if (this.fIndent2) this.indent3 = this.Reader.ReadInt16();
            if (this.fLeftMargin3) this.leftMargin3 = this.Reader.ReadInt16();
            if (this.fIndent3) this.indent3 = this.Reader.ReadInt16();
            if (this.fLeftMargin4) this.leftMargin4 = this.Reader.ReadInt16();
            if (this.fIndent4) this.indent4 = this.Reader.ReadInt16();
            if (this.fLeftMargin5) this.leftMargin5 = this.Reader.ReadInt16();
            if (this.fIndent5) this.indent5 = this.Reader.ReadInt16();

        }        
    }

}
