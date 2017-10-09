using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(4055)]
    public class ExHyperlinkContainer : RegularContainer
    {
        public ExHyperlinkContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecordAttribute(4051)]
    public class ExHyperlinkAtom : Record
    {
        public UInt32 exHyperlinkId;

        public ExHyperlinkAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            exHyperlinkId = this.Reader.ReadUInt32();
        }
    }

    [OfficeRecordAttribute(4068)]
    public class ExHyperlink9Container : RegularContainer
    {
        public ExHyperlink9Container(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }
}
