

using System;
using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(1011)]
    public class SlidePersistAtom : Record
    {
        /// <summary>
        /// logical reference to the slide persist object
        /// </summary>
        public uint PersistIdRef;

        /// <summary>
        /// Bit 1: Slide outline view is collapsed
        /// Bit 2: Slide contains shapes other than placeholders
        /// </summary>
        public uint Flags;

        /// <summary>
        /// number of placeholder texts stored with the persist object.
        /// Allows to display outline view without loading the slide persist objects
        /// </summary>
        public int NumberText;

        /// <summary>
        /// Unique slide identifier, used for OLE link monikers for example
        /// </summary>
        public uint SlideId;

        public SlidePersistAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.PersistIdRef = this.Reader.ReadUInt32();
            this.Flags = this.Reader.ReadUInt32();
            this.NumberText = this.Reader.ReadInt32();
            this.SlideId = this.Reader.ReadUInt32();
            this.Reader.ReadUInt32(); // Throw away reserved field
        }

        override public string ToString(uint depth)
        {

            return string.Format("{0}\n{1}PsrRef = {2}\n{1}Flags = {3}, NumberText = {4}, SlideId = {5})",
                base.ToString(depth), IndentationForDepth(depth + 1),
                this.PersistIdRef, this.Flags, this.NumberText, this.SlideId);
        }
    }
}
