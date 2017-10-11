using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.DocFileFormat
{
    public class BookmarkFirst : ByteStructure
    {
        /// <summary>
        /// An unsigned integer that specifies a zero-based index into the PlcfBkl or PlcfBkld 
        /// that is paired with the PlcfBkf  or PlcfBkfd containing this FBKF. <br/>
        /// The entry found at said index specifies the location of the end of the bookmark associated with this FBKF. <br/>
        /// Ibkl MUST be unique for all FBKFs inside a given PlcfBkf or PlcfBkfd.
        /// </summary>
        public short ibkl;

        /// <summary>
        /// A BKC that specifies further information about the bookmark associated with this FBKF.
        /// </summary>
        public short bkc;

        public BookmarkFirst(VirtualStreamReader reader, int length)
            : base(reader, length)
        {
            this.ibkl = reader.ReadInt16();
            this.bkc = reader.ReadInt16();
        }
    }
}
