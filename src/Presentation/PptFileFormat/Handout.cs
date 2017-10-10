

using System.IO;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecord(4041)]
    public class Handout : RegularContainer
    {
        public Handout(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }

        public SlidePersistAtom PersistAtom;
    }

}
