using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class AnnotationReferenceDescriptor : ByteStructure
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
            var cch = _reader.ReadInt16( );
            var chars = _reader.ReadBytes(18);
            this.UserInitials = Encoding.Unicode.GetString(chars, 0, cch * 2);

            this.AuthorIndex = _reader.ReadUInt16();

            //skip 4 bytes
            _reader.ReadBytes(4);

            this.BookmarkId = _reader.ReadInt32();
        }
    }
}
