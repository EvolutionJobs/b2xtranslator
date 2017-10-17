using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.DocFileFormat
{
    public class FileShapeAddress : ByteStructure
    {
        public enum AnchorType
        {
            margin,
            page,
            text
        }

        /// <summary>
        /// Shape Identifier. Used in conjunction with the office art data 
        /// (found via fcDggInfo in the FIB) to find the actual data for this shape.
        /// </summary>
        public int spid;

        /// <summary>
        /// Left of rectangle enclosing shape relative to the origin of the shape
        /// </summary>
        public int xaLeft;

        /// <summary>
        /// Top of rectangle enclosing shape relative to the origin of the shape
        /// </summary>
        public int yaTop;

        /// <summary>
        /// Right of rectangle enclosing shape relative to the origin of the shape
        /// </summary>
        public int xaRight;

        /// <summary>
        /// Bottom of the rectangle enclosing shape relative to the origin of the shape
        /// </summary>
        public int yaBottom;

        /// <summary>
        /// true in the undo doc when shape is from the header doc<br/>
        /// false otherwise (undefined when not in the undo doc)
        /// </summary>
        public bool fHdr;

        /// <summary>
        /// X position of shape relative to anchor CP<br/>
        /// 0 relative to page margin<br/>
        /// 1 relative to top of page<br/>
        /// 2 relative to text (column for horizontal text; paragraph for vertical text)<br/>
        /// 3 reserved for future use
        /// </summary>
        public AnchorType bx;

        /// <summary>
        /// Y position of shape relative to anchor CP<br/>
        /// 0 relative to page margin<br/>
        /// 1 relative to top of page<br/>
        /// 2 relative to text (column for horizontal text; paragraph for vertical text)<br/>
        /// 3 reserved for future use
        /// </summary>
        public AnchorType by;

        /// <summary>
        /// Text wrapping mode <br/>
        /// 0 like 2, but doesn‘t require absolute object <br/>
        /// 1 no text next to shape <br/>
        /// 2 wrap around absolute object <br/>
        /// 3 wrap as if no object present <br/>
        /// 4 wrap tightly around object <br/>
        /// 5 wrap tightly, but allow holes <br/>
        /// 6-15 reserved for future use
        /// </summary>
        public ushort wr;

        /// <summary>
        /// Text wrapping mode type (valid only for wrapping modes 2 and 4)<br/>
        /// 0 wrap both sides <br/>
        /// 1 wrap only on left <br/>
        /// 2 wrap only on right <br/>
        /// 3 wrap only on largest side
        /// </summary>
        public ushort wrk;

        /// <summary>
        /// When set, temporarily overrides bx, by, 
        /// forcing the xaLeft, xaRight, yaTop, and yaBottom fields 
        /// to all be page relative.
        /// </summary>
        public bool fRcaSimple;

        /// <summary>
        /// true: shape is below text <br/>
        /// false: shape is above text
        /// </summary>
        public bool fBelowText;

        /// <summary>
        /// true: anchor is locked <br/>
        /// fasle: anchor is not locked
        /// </summary>
        public bool fAnchorLock;

        /// <summary>
        /// Count of textboxes in shape (undo doc only)
        /// </summary>
        public int cTxbx;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="reader"></param>
        public FileShapeAddress(VirtualStreamReader reader, int length)
            : base(reader, length)
        {
            this.spid = this._reader.ReadInt32();
            this.xaLeft = this._reader.ReadInt32();
            this.yaTop = this._reader.ReadInt32();
            this.xaRight = this._reader.ReadInt32();
            this.yaBottom = this._reader.ReadInt32();

            ushort flag = this._reader.ReadUInt16();
            this.fHdr = Tools.Utils.BitmaskToBool(flag, 0x0001);
            this.bx = (AnchorType)Tools.Utils.BitmaskToInt(flag, 0x0006);
            this.by = (AnchorType)Tools.Utils.BitmaskToInt(flag, 0x0018);
            this.wr = (ushort)Tools.Utils.BitmaskToInt(flag, 0x01E0);
            this.wrk = (ushort)Tools.Utils.BitmaskToInt(flag, 0x1E00);
            this.fRcaSimple = Tools.Utils.BitmaskToBool(flag, 0x2000);
            this.fBelowText = Tools.Utils.BitmaskToBool(flag, 0x4000);
            this.fAnchorLock = Tools.Utils.BitmaskToBool(flag, 0x8000);

            this.cTxbx = this._reader.ReadInt32();
        }
    }
}
