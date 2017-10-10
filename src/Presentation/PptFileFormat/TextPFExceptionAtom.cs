

using System.IO;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{

    [OfficeRecord(4004)]
    public class TextCFExceptionAtom : Record
    {
        public CharacterRun run;
        public TextCFExceptionAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.run = new CharacterRun(this.Reader);
        }

    }

}
