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
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// TABID: Sheet Tab Index Array (13Dh)
    /// 
    /// This record contains an array of sheet tab index numbers. The record is used by the Shared Lists feature.
    /// 
    /// The sheet tab indexes have type short int (2 bytes each). The index numbers are 0-based 
    /// and are assigned when a sheet is created; the sheets retain their index numbers throughout 
    /// their lifetime in a workbook. If you rearrange the sheets in a workbook, 
    /// the rgiTab array will change to reflect the new sheet arrangement.
    /// 
    /// This record does not appear in BIFF5 files.
    /// </summary>
    [BiffRecordAttribute(RecordType.RRTabId)] 
    public class RRTabId : BiffRecord
    {
        public const RecordType ID = RecordType.RRTabId;

        /// <summary>
        /// Array of tab indexes
        /// 
        /// The index numbers are 0-based and are assigned when a sheet is created; 
        /// the sheets retain their index numbers throughout their lifetime in a workbook. 
        /// If you rearrange the sheets in a workbook, the rgiTab array will 
        /// change to reflect the new sheet arrangement.
        /// </summary>
        public ushort[] rgiTab;

        public RRTabId(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            if (length % 2 != 0)
            {
                throw new RecordParseException(this);
            }
            rgiTab = new ushort[length / 2];
            for (int i = 0; i < length / 2; i++)
            {
                rgiTab[i] = reader.ReadUInt16();
            }
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
