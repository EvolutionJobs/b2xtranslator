

using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{

    [OfficeRecord(4005)]
    public class TextPFExceptionAtom : Record
    {
        public ParagraphRun run;
        public TextPFExceptionAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.Reader.ReadInt16();
            this.run = new ParagraphRun(this.Reader, true);
        }

    }

}
