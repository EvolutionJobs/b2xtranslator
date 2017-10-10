

using System.Text;
using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(4000)]
    public class TextCharsAtom : TextAtom
    {
        public static Encoding ENCODING = Encoding.Unicode;

        public TextCharsAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance, ENCODING) { }
    }

}
