

using b2xtranslator.OfficeDrawing;
using System.IO;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(2032)]
    public class ColorSchemeAtom : Record
    {
        public byte[] Bytes;
        public int Background;
        public int TextAndLines;
        public int Shadows;
        public int TitleText;
        public int Fills;
        public int Accent;
        public int AccentAndHyperlink;
        public int AccentAndFollowedHyperlink;

        public ColorSchemeAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {

                this.Bytes = this.Reader.ReadBytes((int)this.BodySize);
                this.Background = System.BitConverter.ToInt32(this.Bytes, 0);
                this.TextAndLines = System.BitConverter.ToInt32(this.Bytes, 4);
                this.Shadows = System.BitConverter.ToInt32(this.Bytes, 8);
                this.TitleText = System.BitConverter.ToInt32(this.Bytes, 12);
                this.Fills = System.BitConverter.ToInt32(this.Bytes, 16);
                this.Accent = System.BitConverter.ToInt32(this.Bytes, 20);
                this.AccentAndHyperlink = System.BitConverter.ToInt32(this.Bytes, 24);
                this.AccentAndFollowedHyperlink = System.BitConverter.ToInt32(this.Bytes, 28);
        }
    }
}
