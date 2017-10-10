

using System;
using System.IO;

namespace b2xtranslator.PptFileFormat
{
    public class GPointAtom
    {
        public int X;
        public int Y;

        public GPointAtom(BinaryReader reader)
        {
            this.X = reader.ReadInt32();
            this.Y = reader.ReadInt32();
        }

        override public string ToString()
        {
            return string.Format("PointAtom({0}, {1})", this.X, this.Y);
        }
    }

}
