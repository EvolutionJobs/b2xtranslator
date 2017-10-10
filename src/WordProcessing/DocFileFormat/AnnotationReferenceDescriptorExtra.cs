using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class AnnotationReferenceDescriptorExtra : ByteStructure
    {
        public DateAndTime Date;

        public int CommentDepth;

        public int ParentOffset;

        public AnnotationReferenceDescriptorExtra(VirtualStreamReader reader, int length)
            : base(reader, length)
        {
            this.Date = new DateAndTime(_reader.ReadBytes(4));
            _reader.ReadBytes(2);
            this.CommentDepth = _reader.ReadInt32();
            this.ParentOffset = _reader.ReadInt32();
            if (length > 16)
            {
                var flag = _reader.ReadInt32();
            }
        }
    }
}
