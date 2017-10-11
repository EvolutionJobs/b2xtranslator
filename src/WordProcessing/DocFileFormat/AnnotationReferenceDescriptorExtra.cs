using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.DocFileFormat
{
    public class AnnotationReferenceDescriptorExtra : ByteStructure
    {
        public DateAndTime Date;

        public int CommentDepth;

        public int ParentOffset;

        public AnnotationReferenceDescriptorExtra(VirtualStreamReader reader, int length)
            : base(reader, length)
        {
            this.Date = new DateAndTime(this._reader.ReadBytes(4));
            this._reader.ReadBytes(2);
            this.CommentDepth = this._reader.ReadInt32();
            this.ParentOffset = this._reader.ReadInt32();
            if (length > 16)
            {
                int flag = this._reader.ReadInt32();
            }
        }
    }
}
