

using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
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
