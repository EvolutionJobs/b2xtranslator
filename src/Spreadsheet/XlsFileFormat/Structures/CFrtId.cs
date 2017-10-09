using System;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class CFrtId
    {
        /// <summary>
        /// An unsigned integer that specifies the first Future Record Type in the range.<br/> 
        /// MUST be less than or equal to rtLast.
        /// </summary>
        public UInt16 rtFirst;

        /// <summary>
        /// An unsigned integer that specifies the last Future Record Type in the range.
        /// </summary>
        public UInt16 rtLast;

        public CFrtId(IStreamReader reader)
        {
            this.rtFirst = reader.ReadUInt16();
            this.rtLast = reader.ReadUInt16();
        }
    }
}
