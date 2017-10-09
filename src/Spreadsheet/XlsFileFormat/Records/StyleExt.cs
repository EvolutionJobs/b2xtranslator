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
    /// STYLEEXT: Named Cell Style Extension (892h) 
    /// 
    /// This record is used for new Office Excel 2007 formatting properties associated 
    /// with named cell styles. As noted previously XFEXT records are only able to handle 
    /// round-trip formatting when the document was last saved by Office Excel 2007 or later 
    /// and the formatting has not been changed. 
    /// 
    /// This constraint exists because BIFF8 does not have a mechanism for uniquely identifying XF s 
    /// once they are loaded if document formatting changes. For named cell styles however new 
    /// formatting properties can be associated with the style XF by name and the style’s formatting 
    /// can be updated on load (Office Excel 2007 or later).
    /// </summary>
    [BiffRecordAttribute(RecordType.StyleExt)] 
    public class StyleExt : BiffRecord
    {
        public const RecordType ID = RecordType.StyleExt;

        /// <summary>
        /// Record type; this matches the BIFF rt in the first two bytes of the record; =0892h
        /// </summary>
        public UInt16 rt;	

        /// <summary>
        /// FRT cell reference flag; =0 currently
        /// </summary>
        public UInt16 grbitFrt;

        /// <summary>
        /// Currently not used, and set to 0
        /// </summary>
        public UInt64 reserved0;

        /// <summary>
        /// A packed bit field
        /// </summary>
        // private byte grbitFlags; 

        /// <summary>
        /// style category
        /// </summary>
        public byte iCategory;

        /// <summary>
        /// style built in ID 
        /// </summary>
        public byte istyBuiltIn;

        /// <summary>
        /// Level of the outline style RowLevel_n or ColLevel_n. 
        /// 
        /// The automatic outline styles — RowLevel_1 through RowLevel_7, 
        /// and ColLevel_1 through ColLevel_7 — are stored by setting 
        /// istyBuiltIn to 01h or 02h and then setting iLevel to the style level minus 1.  
        /// 
        /// If the style is not an automatic outline style, ignore this field
        /// </summary>
        public byte iLevel;

        /// <summary>
        /// Length of style name (in 2 byte characters)
        /// </summary>
        public UInt16 cchName;

        /// <summary>
        /// Name of style to extend (2 byte characters). If style does not exist then this record is ignored.
        /// </summary>
        public byte[] rgchName;

        /// <summary>
        /// Array of formatting properties. This structure is used to reprsent a set of formatting properties. 
        /// It is described in greater detail in the DXF record description 
        /// </summary>
        //public xfProps	
        // TODO: define class XFPROPS
        
        public StyleExt(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            // TODO: place code here
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
