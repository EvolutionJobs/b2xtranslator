

using System;
using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    public enum PlaceholderEnum
    {
        None = 0,
        MasterTitle = 1,
        MasterBody = 2,
        MasterCenteredTitle = 3,
        MasterSubtitle = 4,
        MasterNotesSlideImage = 5,
        MasterNotesBody = 6,
        MasterDate = 7,
        MasterSlideNumber = 8,
        MasterFooter = 9,
        MasterHeader = 10,
        NotesSlideImage = 11,
        NotesBody = 12,
        Title = 13,
        Body = 14,
        CenteredTitle = 15,
        Subtitle = 16,
        VerticalTextTitle = 17,
        VerticalTextBody = 18,
        Object = 19, // no matter the size
        Graph = 20,
        Table = 21,
        ClipArt = 22,
        OrganizationChart = 23,
        MediaClip = 24
    };

    [OfficeRecord(3011)]
    public class OEPlaceHolderAtom : Record
    {
        /// <summary>
        /// A signed integer that specifies an ID for the placeholder shape.
        /// It SHOULD be unique among all PlaceholderAtom records contained in the corresponding slide.
        /// The value 0xFFFFFFFF specifies that the corresponding shape is not a placeholder shape.
        /// </summary>
        public int Position;

        /// <summary>
        /// A PlaceholderEnum enumeration that specifies the type of the placeholder shape.
        /// </summary>
        public PlaceholderEnum PlacementId;

        /// <summary>
        /// A PlaceholderSize enumeration that specifies the preferred size of the placeholder shape.
        /// </summary>
        public byte PlaceholderSize;

        public OEPlaceHolderAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.Position = this.Reader.ReadInt32();
            this.PlacementId = (PlaceholderEnum) this.Reader.ReadByte();
            this.PlaceholderSize = this.Reader.ReadByte();
            // Throw away additional junk
            this.Reader.ReadUInt16();
        }

        override public string ToString(uint depth)
        {
            return string.Format("{0}\n{1}Position = {2}, PlacementId = {3}, PlaceholderSize = {4}",
                base.ToString(depth), IndentationForDepth(depth + 1),
                this.Position, this.PlacementId, this.PlaceholderSize);
        }

        public bool IsObjectPlaceholder()
        {
            return this.PlacementId == PlaceholderEnum.Object;
        }
    }

}
