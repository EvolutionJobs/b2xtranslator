

using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(1008)]
    public class Note : RegularContainer
    {
        public Note(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }

        public SlidePersistAtom PersistAtom;
    }

}
