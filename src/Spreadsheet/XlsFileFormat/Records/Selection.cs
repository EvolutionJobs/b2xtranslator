/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
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
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecordAttribute(RecordType.Selection)] 
    public class Selection : BiffRecord
    {
        public const RecordType ID = RecordType.Selection;

        /// <summary>
        /// A PaneType that specifies the active pane.
        /// </summary>
        public PaneType pnn;

        /// <summary>
        /// A RwU that specifies the zero-based row number of the active cell.
        /// </summary>
        public ushort rwAct;

        /// <summary>
        /// A ColU that specifies the zero-based column number of the active cell.
        /// </summary>
        public ushort colAct;

        /// <summary>
        /// A signed integer that specifies the zero-based index to the RefU in rgref that 
        /// contains the active cell. MUST be greater than or equal to 0. If this record 
        /// is one of multiple contiguous Selection records, this value is the index to the 
        /// RefU across the aggregation of rgref arrays of all Selection records.
        /// </summary>
        public short irefAct;

        /// <summary>
        /// An unsigned integer that specifies the number of RefU in the rgref array of this record.
        /// </summary>
        public ushort cref;

        /// <summary>
        /// An array of RefU that specifies ranges of selected cells in the sheet. 
        /// The number of RefU in this array MUST be less than or equal to 1369.
        /// </summary>
        public RefU[] rgref;

        public Selection(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.pnn = (PaneType)reader.ReadByte();
            this.rwAct = reader.ReadUInt16();
            this.colAct = reader.ReadUInt16();
            this.irefAct = reader.ReadInt16();
            this.cref = reader.ReadUInt16();

            if (this.cref > 0)
            {
                this.rgref = new RefU[this.cref];
                for (int i = 0; i < this.cref; i++)
                {
                    this.rgref[i] = new RefU(reader);
                }
            }
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
