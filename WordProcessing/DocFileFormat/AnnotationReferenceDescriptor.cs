using System.Text;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.DocFileFormat
{
    public sealed class AnnotationReferenceDescriptor : ByteStructure
    {
        /// <summary>
        /// The initials of the user who left the annotation.
        /// </summary>
        public string UserInitials;

        /// <summary>
        /// An index into the string table of comment author names.
        /// </summary>
        public ushort AuthorIndex;

        /// <summary>
        /// Identifies a bookmark.
        /// </summary>
        public int BookmarkId;

        public AnnotationReferenceDescriptor(VirtualStreamReader reader, int length)
            : base(reader, length)
        {
            //read the user initials (LPXCharBuffer9)
            short cch = this._reader.ReadInt16( );
            var chars = this._reader.ReadBytes(18);
            this.UserInitials = Encoding.Unicode.GetString(chars, 0, cch * 2);

            this.AuthorIndex = this._reader.ReadUInt16();

            //skip 4 bytes
            this._reader.ReadBytes(4);

            this.BookmarkId = this._reader.ReadInt32();
        }
    }
}
