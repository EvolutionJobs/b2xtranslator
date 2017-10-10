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
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// FORMAT: Number Format (41Eh)
    /// 
    /// The FORMAT record describes a number format in the workbook.
    /// 
    /// All the FORMAT records should appear together in a BIFF file. 
    /// The order of FORMAT records in an existing BIFF file should not be changed.  
    /// It is possible to write custom number formats in a file, but they 
    /// should be added at the end of the existing FORMAT records.
    /// </summary>
    [BiffRecordAttribute(RecordType.Format)] 
    public class Format : BiffRecord
    {
        public const RecordType ID = RecordType.Format;
        
        /// <summary>
        /// Format index code (for internal use only)
        /// 
        /// Excel uses the ifmt field to identify built-in formats when it reads a file 
        /// that was created by a different localized version. 
        /// 
        /// For more information about built-in formats, see <code>XF</code> Extended Format (E0h). 
        /// </summary>
        public ushort ifmt;	

        /// <summary>
        /// Length of the string
        /// </summary>
        public ushort cch;	

        /// <summary>
        /// Option Flags (described in Unicode Strings in BIFF8 section) 
        /// </summary>
        public byte grbit;	

        /// <summary>
        /// Array of string characters
        /// </summary>
        public string rgb;	


        public Format(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            ifmt = reader.ReadUInt16();
            cch = reader.ReadUInt16();
            grbit = reader.ReadByte();
            // TODO: place code for interpretation of grbit flag here
            // TODO: possibly define a wrapper class for Unicode strings

            rgb = ExcelHelperClass.getStringFromBiffRecord(reader, cch, grbit); 
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
