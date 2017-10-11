

using System;
using System.IO;

namespace b2xtranslator.PptFileFormat
{
    public class GRatioAtom
    {
        public int Numer;
        public int Denom;

        public GRatioAtom(BinaryReader reader)
        {
            this.Numer = reader.ReadInt32();
            this.Denom = reader.ReadInt32();
        }

        override public string ToString()
        {
            return string.Format("RatioAtom({0}, {1})", this.Numer, this.Denom);
        }
    }

}
