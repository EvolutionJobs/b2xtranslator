using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Drawing;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    [OfficeRecord(0xF00F)]
    public class ChildAnchor : Record
    {
        /// <summary>
        /// Rectangle that describe sthe bounds of the anchor
        /// </summary>
        public Rectangle rcgBounds;
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;

        public ChildAnchor(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            Left = this.Reader.ReadInt32();
            Top = this.Reader.ReadInt32();
            Right = this.Reader.ReadInt32();
            Bottom = this.Reader.ReadInt32();
            this.rcgBounds = new Rectangle(
                new Point(Left, Top),
                new Size((Right-Left), (Bottom-Top))
            );
        }
    }
}
