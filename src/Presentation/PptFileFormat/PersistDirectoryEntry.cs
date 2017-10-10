/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
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
            var StartPersistIdAndPersistCount = reader.ReadUInt32();
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

            foreach (var entry in this.PersistOffsetEntries)
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
