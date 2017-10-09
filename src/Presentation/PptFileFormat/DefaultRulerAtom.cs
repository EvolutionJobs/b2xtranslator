using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.IO;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(4011)]
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
            flags = Reader.ReadUInt32();
            fDefaultTabSize = Utils.BitmaskToBool(flags, 0x1);
            fCLevels = Utils.BitmaskToBool(flags, 0x1 << 1);
            fTabStops = Utils.BitmaskToBool(flags, 0x1 << 2);
            fLeftMargin1 = Utils.BitmaskToBool(flags, 0x1 << 3);
            fLeftMargin2 = Utils.BitmaskToBool(flags, 0x1 << 4);
            fLeftMargin3 = Utils.BitmaskToBool(flags, 0x1 << 5);
            fLeftMargin4 = Utils.BitmaskToBool(flags, 0x1 << 6);
            fLeftMargin5 = Utils.BitmaskToBool(flags, 0x1 << 7);
            fIndent1 = Utils.BitmaskToBool(flags, 0x1 << 8);
            fIndent2 = Utils.BitmaskToBool(flags, 0x1 << 9);
            fIndent3 = Utils.BitmaskToBool(flags, 0x1 << 10);
            fIndent4 = Utils.BitmaskToBool(flags, 0x1 << 11);
            fIndent5 = Utils.BitmaskToBool(flags, 0x1 << 12);

            if (fCLevels) cLevels = Reader.ReadInt16();
            if (fDefaultTabSize) defaultTabSize = Reader.ReadInt16();

            if (fTabStops) tabs = new TabStops(Reader);

            if (fLeftMargin1) leftMargin1 = Reader.ReadInt16();
            if (fIndent1) indent1 = Reader.ReadInt16();
            if (fLeftMargin2) leftMargin2 = Reader.ReadInt16();
            if (fIndent2) indent3 = Reader.ReadInt16();
            if (fLeftMargin3) leftMargin3 = Reader.ReadInt16();
            if (fIndent3) indent3 = Reader.ReadInt16();
            if (fLeftMargin4) leftMargin4 = Reader.ReadInt16();
            if (fIndent4) indent4 = Reader.ReadInt16();
            if (fLeftMargin5) leftMargin5 = Reader.ReadInt16();
            if (fIndent5) indent5 = Reader.ReadInt16();

        }        
    }

}
