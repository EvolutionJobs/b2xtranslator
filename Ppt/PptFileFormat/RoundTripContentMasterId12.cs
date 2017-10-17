

using System;
using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(1058)]
    public class RoundTripContentMasterId12 : Record
    {
        /// <summary>
        /// Round-trip id of the main master
        /// </summary>
        public uint MainMasterId;

        /// <summary>
        /// Instance id of the content master (unique for main master)
        /// </summary>
        public uint ContentMasterInstanceId;

        public RoundTripContentMasterId12(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.MainMasterId = this.Reader.ReadUInt32();
            this.ContentMasterInstanceId = this.Reader.ReadUInt32();
        }

        override public string ToString(uint depth)
        {
            return string.Format("{0}\n{1}MainMasterId = {2}, ContentMasterInstanceId = {3}",
                base.ToString(depth), IndentationForDepth(depth + 1),
                this.MainMasterId, this.ContentMasterInstanceId);
        }
    }

}
