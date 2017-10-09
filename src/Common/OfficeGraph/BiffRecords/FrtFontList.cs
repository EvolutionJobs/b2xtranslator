/*
 * Copyright (c) 2009, DIaLOGIKa
 *
 * All rights reserved.
 * 
 * Redistribution and use in source and binary forms, with or without 
 * modification, are permitted provided that the following conditions are met:
 * 
 *     * Redistributions of source code must retain the above copyright 
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright 
 *       notice, this list of conditions and the following disclaimer in the 
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the names of copyright holders, nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software 
 *       without specific prior written permission.
 * 
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" 
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED 
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. 
 * IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, 
 * INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, 
 * BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, 
 * DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF 
 * LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE 
 * OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF 
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE. 
 */

using System;
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies font information used on the chart and specifies 
    /// the beginning of a collection of Font records as defined by the Chart Sheet Substream ABNF.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.FrtFontList)]
    public class FrtFontList : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.FrtFontList;

        /// <summary>
        /// An FrtHeaderOld. The frtHeaderOld.rt field MUST be 0x085A.
        /// </summary>
        FrtHeaderOld frtHeaderOld;

        /// <summary>
        /// An unsigned integer that specifies the application version where new chart elements
        /// were introduced that use the font information specified by rgFontInfo. 
        /// 
        /// MUST be equal to iObjectInstance1 of the StartObject record that immediately 
        /// follows this record as defined by the Chart Sheet Substream ABNF and 
        /// MUST be a value from the following table: 
        /// 
        ///     Value       Meaning
        ///     0x09        This record pertains to new objects introduced in a specific 
        ///                 version of the application <53>. rgFontInfo specifies the font 
        ///                 information that is used by display units labels specified by YMult.
        ///                 
        ///     0x0A        This record pertains to new objects introduced in specific version of 
        ///                 the application <54>. rgFontInfo specifies the font information that 
        ///                 is used by extended data labels specified by DataLabExt.
        /// </summary>
        public byte verExcel;

        /// <summary>
        /// An unsigned integer that specifies the number of items in rgFontInfo.
        /// </summary>
        public UInt16 cFont;

        /// <summary>
        /// An array of FontInfo structures that specify the font information. 
        /// 
        /// The number of elements in this array MUST be equal to the value specified in cFont.
        /// </summary>
        public FontInfo[] rgFontInfo;
        
        public FrtFontList(IStreamReader reader, GraphRecordNumber id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.frtHeaderOld = new FrtHeaderOld(reader);
            this.verExcel = reader.ReadByte();

            // skip reserved byte
            reader.ReadByte();

            this.cFont = reader.ReadUInt16();

            if (this.cFont > 0)
            {
                this.rgFontInfo = new FontInfo[this.cFont];

                for (int i = 0; i < this.cFont; i++)
                {
                    this.rgFontInfo[i] = new FontInfo(reader);                    
                }
            }

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
