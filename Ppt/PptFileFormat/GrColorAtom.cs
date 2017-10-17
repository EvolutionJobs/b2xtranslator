

using System;
using System.IO;

namespace b2xtranslator.PptFileFormat
{
    public class GrColorAtom
    {
        public byte Red;
        public byte Green;
        public byte Blue;
        public byte Index;

        public bool IsSchemeColor
        {
            get { return this.Index != 0xFE; }
        }

        public GrColorAtom(BinaryReader reader)
        {
            this.Red = reader.ReadByte();
            this.Green = reader.ReadByte();
            this.Blue = reader.ReadByte();
            this.Index = reader.ReadByte();
        }

        public override string ToString()
        {
            return string.Format("GrColorAtom({0}, {1}, {2}): Index = {3}",
                this.Red, this.Green, this.Blue, this.Index);
        }
    }

}
