

using System;
using System.Text;
using System.IO;
using b2xtranslator.OfficeDrawing;
using b2xtranslator.Tools;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(4026)]
    public class CStringAtom : Record
    {
        public static Encoding ENCODING = Encoding.Unicode;
        public string Text;

        public CStringAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            var bytes = new byte[size];
            this.Reader.Read(bytes, 0, (int)size);

            this.Text = new string(ENCODING.GetChars(bytes));
        }

        public override string ToString(uint depth)
        {
            return string.Format("{0}\n{1}Text = {2}",
                base.ToString(depth), IndentationForDepth(depth + 1), Utils.StringInspect(this.Text));
        }
    }
}
