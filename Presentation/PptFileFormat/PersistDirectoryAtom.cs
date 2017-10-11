

using System.Collections.Generic;
using System.Text;
using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(6002)]
    public class PersistDirectoryAtom : Record
    {
        /// <summary>
        /// An array of PersistDirectoryEntry structures that specifies persist 
        /// object identifiers and stream offsets to persist objects.
        /// 
        /// The size, in bytes, of the array is specified by rh.recLen.  
        /// </summary>
        public List<PersistDirectoryEntry> PersistDirEntries = new List<PersistDirectoryEntry>();

        public PersistDirectoryAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            while (this.Reader.BaseStream.Position != this.Reader.BaseStream.Length)
            {
                this.PersistDirEntries.Add(new PersistDirectoryEntry(this.Reader));
            }
        }

        override public string ToString(uint depth)
        {
            var sb = new StringBuilder();

            sb.Append(base.ToString(depth));

            foreach (var entry in this.PersistDirEntries)
            {
                sb.AppendLine();
                sb.Append(entry.ToString(depth + 1));
            }

            return sb.ToString();
        }
    }

}
