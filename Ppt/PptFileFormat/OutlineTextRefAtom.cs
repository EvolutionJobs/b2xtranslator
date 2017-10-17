

using System;
using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    /// <summary>
    /// An atom record that specifies a reference to text contained in the SlideListWithTextContainer record. 
    /// 
    /// Let the corresponding slide persist be specified by the SlidePersistAtom record contained in the 
    /// SlideListWithTextContainer record whose persistIdRef field refers to the SlideContainer record that 
    /// contains this OutlineTextRefAtom record. 
    /// </summary>
    [OfficeRecord(3998)]
    public class OutlineTextRefAtom : Record
    {
        /// <summary>
        /// A signed integer that specifies a zero-based index into the sequence of 
        /// TextHeaderAtom records that follows the corresponding slide persist.
        /// 
        /// It MUST be greater than or equal to 0x00000000 and less than the number
        /// of TextHeaderAtom records that follow the corresponding slide persist. 
        /// </summary>
        public int Index;

        public OutlineTextRefAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.Index = this.Reader.ReadInt32();
        }

        override public string ToString(uint depth)
        {
            return string.Format("{0}\n{1}Index = {2}",
                base.ToString(depth), IndentationForDepth(depth + 1),
                this.Index);
        }
    }

}
