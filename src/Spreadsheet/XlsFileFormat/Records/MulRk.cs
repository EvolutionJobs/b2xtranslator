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
using System.Collections.Generic;
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecordAttribute(RecordType.MulRk)] 
    public class MulRk : BiffRecord
    {
        public const RecordType ID = RecordType.MulRk;

        /// <summary>
        /// Row 
        /// </summary>
        public UInt16 rw;

        /// <summary>
        /// First column number 
        /// </summary>
        public UInt16 colFirst;        

        /// <summary>
        /// The last affected column 
        /// </summary>
        public UInt16 colLast;         

        /// <summary>
        /// List with format indexes 
        /// </summary>
        public List<UInt16> ixfe;      // List records 

        /// <summary>
        /// List with the numbers 
        /// </summary>
        public List<double> rknumber;  // List double values 


        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader">Streamreader</param>
        /// <param name="id">Record ID - Recordtype</param>
        /// <param name="length">The recordlegth</param>
        public MulRk(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);
            this.ixfe = new List<UInt16>();
            this.rknumber = new List<double>(); 

            // count records - 6 standard non variable values !!! 
            int count = (int)(this.Length - 6) / 6 ;
            this.rw = reader.ReadUInt16();
            this.colFirst = reader.ReadUInt16();
            for (int i = 0; i < count; i++)
            {
                this.ixfe.Add(reader.ReadUInt16());
                Byte[] buffer = reader.ReadBytes(4);

                rknumber.Add(ExcelHelperClass.NumFromRK(buffer)); 
            }
            this.colLast = reader.ReadUInt16(); 
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
       }
    }
}
