

using System.Collections.Generic;
using System.Text;
using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    /// <summary>
    /// A structure that specifies a compressed table of sequential persist
    /// object identifiers and stream offsets to associated persist objects. 
    /// 
    /// Let the corresponding user edit be specified by the UserEditAtom record that most closely follows the 
    /// PersistDirectoryAtom record that contains this structure. 
    /// 
    /// Let the corresponding persist object directory be specified by the PersistDirectoryAtom record that 
    /// contains this structure. 
    /// </summary>
    public class PersistDirectoryEntry
    {
        /// <summary>
        /// An unsigned integer that specifies a starting persist object identifier. (20 bits)
        /// 
        /// It MUST be less than or equal to 0xFFFFE.
        /// 
        /// The first entry in PersistOffsetEntries is associated with StartPersistId.
        /// 
        /// The next entry, if present, is associated with StartPersistId + 1.
        /// 
        /// Each entry in PersistOffsetEntries is associated with a persist object identifier in this manner,
        /// with the final entry associated with StartPersistId + PersistCount – 1. 
        /// </summary>
        public uint StartPersistId;

        /// <summary>
        /// An unsigned integer that specifies the count of items in PersistOffsetEntries. (12 bit)
        /// It MUST be greater than or equal to 0x001. 
        /// </summary>
        public uint PersistCount;

        /// <summary>
        /// An array of PersistOffsetEntry that specifies stream offsets to persist 
        /// objects. The count of items in the array is specified by PersistCount.
        /// 
        /// The value of each item MUST be greater than or equal to OffsetLastEdit
        /// in the corresponding user edit and MUST be less than the offset, in bytes,
        /// of the corresponding persist object directory. 
        /// </summary>
        public List<uint> PersistOffsetEntries = new List<uint>();

        public PersistDirectoryEntry(BinaryReader reader)
        {
            uint StartPersistIdAndPersistCount = reader.ReadUInt32();
            this.StartPersistId = (StartPersistIdAndPersistCount & 0x000FFFFFU); // First 20 bit
            this.PersistCount   = (StartPersistIdAndPersistCount & 0xFFF00000U) >> 20; // Last 12 bit

            for (int i = 0; i < this.PersistCount; i++)
            {
                this.PersistOffsetEntries.Add(reader.ReadUInt32());
            }
        }

        public string ToString(uint depth)
        {
            var sb = new StringBuilder();

            sb.AppendFormat("{0}PersistDirectoryEntry: ", Record.IndentationForDepth(depth));

            bool isFirst = true;

            foreach (uint entry in this.PersistOffsetEntries)
            {
                if (!isFirst)
                    sb.Append(", ");

                sb.Append(entry);

                isFirst = false;
            }

            return sb.ToString();
        }
    }
}
