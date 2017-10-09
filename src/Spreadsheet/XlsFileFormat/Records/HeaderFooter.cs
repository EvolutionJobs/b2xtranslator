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
    [BiffRecordAttribute(RecordType.HeaderFooter)] 
    public class HeaderFooter : BiffRecord
    {
        public const RecordType ID = RecordType.HeaderFooter;

        /// <summary>
        /// An FrtHeader. The frtHeader.rt field MUST be 0x089C.
        /// </summary>
        public FrtHeader frtHeader;

        /// <summary>
        /// A GUID as specified by [MS-DTYP] that specifies the current sheet view. <br/>
        /// If it is zero, it means the current sheet. <br/>
        /// Otherwise, this field MUST match the guid field of the preceding UserSViewBegin record.
        /// </summary>
        public Guid guidSView;

        /// <summary>
        /// An unsigned integer that specifies the number of characters in strHeaderEven. <br/>
        /// MUST be less than or equal to 255. <br/>
        /// MUST be zero if fHFDiffOddEven is zero.
        /// </summary>
        public UInt16 cchHeaderEven;

        /// <summary>
        /// An unsigned integer that specifies the number of characters in strFooterEven.<br/> 
        /// MUST be less than or equal to 255. <br/>
        /// MUST be zero if fHFDiffOddEven is zero.
        /// </summary>
        public UInt16 cchFooterEven;

        /// <summary>
        /// An unsigned integer that specifies the number of characters in strHeaderFirst.<br/> 
        /// MUST be less than or equal to 255. <br/>
        /// MUST be zero if fHFDiffFirst is zero.
        /// </summary>
        public UInt16 cchHeaderFirst;

        /// <summary>
        /// An unsigned integer that specifies the number of characters in strFooterFirst.<br/>
        /// MUST be less than or equal to 255. <br/>
        /// MUST be zero if fHFDiffFirst is zero.
        /// </summary>
        public UInt16 cchFooterFirst;

        /// <summary>
        /// An XLUnicodeString that specifies the header text on the even pages. <br/>
        /// The number of characters in the string MUST be equal to cchHeaderEven. <br/>
        /// The string can contain special commands, for example a placeholder for the page number, 
        /// current date or text formatting attributes. <br/>
        /// Refer to Header for more details about the string format.
        /// </summary>
        public string strHeaderEven;

        /// <summary>
        /// An XLUnicodeString that specifies the footer text on the even pages.<br/> 
        /// The number of characters in the string MUST be equal to cchFooterEven.<br/> 
        /// The string can contain special commands, for example a placeholder for the page number, 
        /// current date or text formatting attributes.<br/> 
        /// Refer to Header for more details about the string format.
        /// </summary>
        public string strFooterEven;

        /// <summary>
        /// An XLUnicodeString that specifies the header text on the first page. <br/> 
        /// The number of characters in the string MUST be equal to cchHeaderFirst. <br/> 
        /// The string can containspecial commands, for example a placeholder for the page number, 
        /// current date or text formatting attributes. <br/> 
        /// Refer to Header for more details about the string format.
        /// </summary>
        public string strHeaderFirst;

        /// <summary>
        /// An XLUnicodeString that specifies the footer text on the first page. <br/> 
        /// The number of characters in the string MUST be equal to cchFooterFirst. <br/> 
        /// The string can contain special commands, for example a placeholder for the page number, 
        /// current date or text formatting attributes. <br/> 
        /// Refer to Header for more details about the string format.
        /// </summary>
        public string strFooterFirst;

        public HeaderFooter(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.frtHeader = new FrtHeader(reader);
            this.guidSView = new Guid(reader.ReadBytes(16));
            UInt16 flags = reader.ReadUInt16();
            this.cchHeaderEven = reader.ReadUInt16();
            this.cchFooterEven = reader.ReadUInt16();
            this.cchHeaderFirst = reader.ReadUInt16();
            this.cchFooterFirst = reader.ReadUInt16();

            byte[] strHeaderEvenBytes = reader.ReadBytes(cchHeaderEven);
            byte[] strFooterEvenBytes = reader.ReadBytes(cchFooterEven);
            byte[] strHeaderFirstBytes = reader.ReadBytes(cchHeaderFirst);
            byte[] strFooterFirstBytes = reader.ReadBytes(cchFooterFirst);

            //this.strHeaderEven = new XLUnicodeString(reader).Value;
            //this.strFooterEven = new XLUnicodeString(reader).Value;
            //this.strHeaderFirst = new XLUnicodeString(reader).Value;
            //this.strFooterFirst = new XLUnicodeString(reader).Value;

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
