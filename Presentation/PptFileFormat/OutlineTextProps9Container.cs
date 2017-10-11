using System.Collections.Generic;
using b2xtranslator.OfficeDrawing;
using System.IO;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(4014)]
    public class OutlineTextProps9Container : RegularContainer
    {
        public List<OutlineTextProps9Entry> OutlineTextProps9Entries = new List<OutlineTextProps9Entry>();

        public OutlineTextProps9Container(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {

            this.Reader.BaseStream.Position = 0;
                while (this.Reader.BaseStream.Position < this.Reader.BaseStream.Length)
                {
                    var entry = new OutlineTextProps9Entry(this.Reader);
                this.OutlineTextProps9Entries.Add(entry);
                }
        }
    }

    [OfficeRecord(4015)]
    public class OutlineTextPropsHeader9Atom : Record
    {
        public int slideIdRef;
        public int txType;

        public OutlineTextPropsHeader9Atom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.slideIdRef = this.Reader.ReadInt32();
            this.txType = this.Reader.ReadInt32();
        }
    }


    public class OutlineTextProps9Entry
    {
        public OutlineTextPropsHeader9Atom outlineTextHeaderAtom;
        public StyleTextProp9Atom styleTextProp9Atom;

        public OutlineTextProps9Entry(BinaryReader reader)
        {
            this.outlineTextHeaderAtom = (OutlineTextPropsHeader9Atom)Record.ReadRecord(reader);
            this.styleTextProp9Atom = (StyleTextProp9Atom)Record.ReadRecord(reader);
        }
    }
}
