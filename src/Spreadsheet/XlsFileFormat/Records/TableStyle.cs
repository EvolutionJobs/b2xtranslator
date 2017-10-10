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
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// TABLESTYLE: Table Style (88Fh)
    /// 
    /// This record is used for each custom Table style in use in the document.
    /// </summary>
    [BiffRecordAttribute(RecordType.TableStyle)] 
    public class TableStyle : BiffRecord
    {
        public const RecordType ID = RecordType.TableStyle;

        /// <summary>
        /// Record type; this matches the BIFF rt in the first two bytes of the record; =088Fh
        /// </summary>
        public ushort rt;

        /// <summary>
        /// FRT cell reference flag; =0 currently
        /// </summary>
        public ushort grbitFrt;

        /// <summary>
        /// Currently not used, and set to 0
        /// </summary>
        public UInt64 reserved0;

        /// <summary>
        /// A packed bit field
        /// </summary>
        private ushort grbitTS;

        /// <summary>
        /// Count of TABLESTYLEELEMENT records to follow.
        /// </summary>
        public UInt32 ctse;

        /// <summary>
        /// Length of Table style name in 2 byte characters.
        /// </summary>
        public ushort cchName;

        /// <summary>
        /// Table style name in 2 byte characters
        /// </summary>
        public byte[] rgchName;

        /// <summary>
        /// Should always be 0.
        /// </summary>
        public bool fIsBuiltIn;

        /// <summary>
        /// =1 if Table style can be applied to PivotTables
        /// </summary>
        public bool fIsPivot;

        /// <summary>
        /// =1 if Table style can be applied to Tables
        /// </summary>
        public bool fIsTable;

        /// <summary>
        /// Reserved; must be 0 (zero)
        /// </summary>
        public ushort fReserved0;


        public TableStyle(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            rt = reader.ReadUInt16();
            grbitFrt = reader.ReadUInt16();
            reserved0 = reader.ReadUInt64();
            grbitTS = reader.ReadUInt16();

            fIsBuiltIn = Utils.BitmaskToBool(grbitTS, 0x0001);
            fIsPivot = Utils.BitmaskToBool(grbitTS, 0x0002);
            fIsTable = Utils.BitmaskToBool(grbitTS, 0x0004);
            fReserved0 = (ushort)Utils.BitmaskToInt(grbitTS, 0xFFF8);
            
            ctse = reader.ReadUInt32();
            cchName = reader.ReadUInt16();
            rgchName = reader.ReadBytes(cchName * 2);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
