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
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecordAttribute(RecordType.Array)] 
    public class ARRAY : BiffRecord
    {
        public const RecordType ID = RecordType.Array;
        /// <summary>
        /// Rownumber 
        /// </summary>
        public ushort rwFirst;

        /// <summary>
        /// Rownumber 
        /// </summary>
        public ushort rwLast;

        /// <summary>
        /// Colnumber 
        /// </summary>
        public byte colFirst;

        /// <summary>
        /// Colnumber 
        /// </summary>
        public byte colLast;

        /// <summary>
        /// option flags 
        /// </summary>
        public ushort grbit;

        /// <summary>
        /// used for performance reasons only 
        /// can be ignored 
        /// </summary>
        public uint chn;

        /// <summary>
        /// length of the formular data !!
        /// </summary>
        public ushort cce;

        /// <summary>
        /// LinkedList with the Ptg records !!
        /// </summary>
        public Stack<AbstractPtg> ptgStack;

        public ARRAY(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.rwFirst = reader.ReadUInt16();
            this.rwLast = reader.ReadUInt16();
            this.colFirst = reader.ReadByte();
            this.colLast = reader.ReadByte();
            this.grbit = reader.ReadUInt16();
            this.chn = reader.ReadUInt32(); // this is used for performance reasons only 
            this.cce = reader.ReadUInt16();
            this.ptgStack = new Stack<AbstractPtg>();
            // reader.ReadBytes(this.cce);

            this.ptgStack = ExcelHelperClass.getFormulaStack(this.Reader, this.cce); 
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
