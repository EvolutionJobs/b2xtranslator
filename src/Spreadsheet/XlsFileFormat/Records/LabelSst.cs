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
    /// This class extracts the data from a LABELSST Record
    /// This record describes a cell that contains a string constant from the shared string table 
    /// </summary>
    [BiffRecordAttribute(RecordType.LabelSst)] 
    public class LabelSst : BiffRecord
    {
        public const RecordType ID = RecordType.LabelSst;

        /// <summary>
        /// Some public attributes to store the data from this record 
        /// </summary>

        /// <summary>
        /// Rownumber 
        /// </summary>
        public UInt16 rw;      

        /// <summary>
        /// Colnumber 
        /// </summary>
        public UInt16 col;     
        
        /// <summary>
        /// index to the XF record 
        /// </summary>
        public UInt16 ixfe;    

        /// <summary>
        /// index into the SST record  
        /// </summary>
        public UInt32 isst;     

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="id"></param>
        /// <param name="length"></param>
        public LabelSst(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);
            this.rw = reader.ReadUInt16();
            this.col = reader.ReadUInt16();
            this.ixfe = reader.ReadUInt16();
            this.isst = reader.ReadUInt32();
           
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
