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
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies the settings of a Page Layout view for a sheet.
    /// </summary>
    [BiffRecordAttribute(RecordType.PLV)] 
    public class PLV : BiffRecord
    {
        public const RecordType ID = RecordType.PLV;

        /// <summary>
        /// An FrtHeader. The frtHeader.rt field MUST be 0x088B.
        /// </summary>
        public FrtHeader frtHeader;

        /// <summary>
        /// An unsigned integer that specifies zoom scale as a percentage for the 
        /// Page Layout view of the current sheet. For example, if the value is 107, 
        /// then the zoom scale is 107%. The value 0 means that the zoom scale has not 
        /// been set. If the value is nonzero, it MUST be greater than or equal 
        /// to 10 and less than or equal to 400.
        /// </summary>
        public UInt16 wScalePLV;

        /// <summary>
        /// A bit that specifies whether the sheet is in the Page Layout view. 
        /// If the fSLV in Window2 record is 1 for this sheet, it MUST be 0
        /// </summary>
        public bool fPageLayoutView;

        /// <summary>
        /// A bit that specifies whether the application displays the ruler.
        /// </summary>
        public bool fRulerVisible;

        /// <summary>
        /// A bit that specifies whether the margins between pages are hidden in the Page Layout view.
        /// </summary>
        public bool fWhitespaceHidden;


        public PLV(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.frtHeader = new FrtHeader(reader);
            this.wScalePLV = reader.ReadUInt16();

            UInt16 flags = reader.ReadUInt16();
            this.fPageLayoutView = Utils.BitmaskToBool(flags, 0x0001);
            this.fRulerVisible = Utils.BitmaskToBool(flags, 0x0002);
            this.fWhitespaceHidden = Utils.BitmaskToBool(flags, 0x0004);
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
