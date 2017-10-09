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

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies the text elements that are formatted using the 
    /// information specified by the Text record immediately following this record.
    /// </summary>
    [BiffRecordAttribute(RecordType.DefaultText)]
    public class DefaultText : BiffRecord
    {
        public const RecordType ID = RecordType.DefaultText;

        public enum DefaultTextType : ushort
        {
            NoShowPercentNoShowValue = 0x0000,
            ShowPercentNoShowValue = 0x0001,
            ScalableNoFontInfo = 0x0002,
            ScalableFontInfo = 0x0003
        }

        /// <summary>
        /// An unsigned integer that specifies the text elements that are formatted using 
        /// the position and appearance information specified by the Text record
        /// immediately following this record.
        /// 
        /// If this record is located in the sequence of records that conforms to the CRT 
        /// rule as specified by the Chart Sheet Substream ABNF, this record MUST be 0x0000 or 0x0001. 
        /// 
        /// If this record is not located in the CRT rule as specified by the Chart Sheet 
        /// Substream ABNF, this record MUST be 0x0002 or 0x0003. 
        /// 
        /// MUST be a value from the following table: 
        /// 
        ///     Value       Meaning
        ///     0x0000      Format all Text records in the chart group where fShowPercent equals 0 or fShowValue equals 0.
        ///     0x0001      Format all Text records in the chart group where fShowPercent equals 1 or fShowValue equals 1.
        ///     0x0002      Format all Text records in the chart where the value of fScalable of the associated FontInfo structure equals 0.
        ///     0x0003      Format all Text records in the chart where the value of fScalable of the associated FontInfo structure equals 1.
        /// </summary>
        public DefaultTextType defaultTextId;

        public DefaultText(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.defaultTextId = (DefaultTextType)reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
