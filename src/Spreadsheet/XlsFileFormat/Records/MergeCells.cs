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
using System.Collections.Generic;
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This class extracts the data from a mergecell biffrecord 
    /// 
    /// MERGECELLS: Merged Cells (E5h) 
    /// This record stores all merged cells. Record Data 
    /// Offset 	Field Name 	Size 	Contents 
    /// 4 	cmcs 	2 	Count of REF structures
    /// 
    /// The REF structure has the following fields. 
    /// Offset 	Field Name 	Size 	Contents 
    /// 8 	rwFirst 	2 	The first row of the range associated with the record 
    /// 10 	rwLast 	2 	The last row of the range associated with the record 
    /// 12 	colFirst 	2 	The first column of the range associated with the record 
    /// 14 	colLast 	2 	The last column of the range associated with the record 
    /// </summary>
    [BiffRecordAttribute(RecordType.MergeCells)] 
    public class MergeCells : BiffRecord
    {
        public const RecordType ID = RecordType.MergeCells;

        /// <summary>
        /// List with datarecords 
        /// </summary>
        public List<MergeCellData> mergeCellDataList;

        /// <summary>
        /// Count REF structures 
        /// </summary>
        public ushort cmcs;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="id"></param>
        /// <param name="length"></param>
        public MergeCells(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            this.mergeCellDataList = new List<MergeCellData>(); 
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.cmcs = this.Reader.ReadUInt16();

            for (int i = 0; i < this.cmcs; i++)
            {
                var mcd = new MergeCellData();
                mcd.rwFirst = this.Reader.ReadUInt16();
                mcd.rwLast = this.Reader.ReadUInt16();
                mcd.colFirst = this.Reader.ReadUInt16();
                mcd.colLast = this.Reader.ReadUInt16();
                this.mergeCellDataList.Add(mcd); 
            }
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
