using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecord(4014)]
    public class OutlineTextProps9Container : RegularContainer
    {
        public List<OutlineTextProps9Entry> OutlineTextProps9Entries = new List<OutlineTextProps9Entry>();

        public OutlineTextProps9Container(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {

                Reader.BaseStream.Position = 0;
                while (Reader.BaseStream.Position < Reader.BaseStream.Length)
                {
                    var entry = new OutlineTextProps9Entry(Reader);
                    OutlineTextProps9Entries.Add(entry);
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
            slideIdRef = Reader.ReadInt32();
            txType = Reader.ReadInt32();
        }
    }


    public class OutlineTextProps9Entry
    {
        public OutlineTextPropsHeader9Atom outlineTextHeaderAtom;
        public StyleTextProp9Atom styleTextProp9Atom;

        public OutlineTextProps9Entry(BinaryReader reader)
        {
            outlineTextHeaderAtom = (OutlineTextPropsHeader9Atom)Record.ReadRecord(reader);
            styleTextProp9Atom = (StyleTextProp9Atom)Record.ReadRecord(reader);
        }
    }
}
