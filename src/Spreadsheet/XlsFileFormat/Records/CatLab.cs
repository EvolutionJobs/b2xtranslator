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
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies the attributes of the axis label.
    /// </summary>
    [BiffRecordAttribute(RecordType.CatLab)]
    public class CatLab : BiffRecord
    {
        public const RecordType ID = RecordType.CatLab;

        public enum Alignment : ushort
        {
            /// <summary>
            /// Top-aligned if the trot field of the Text record of the axis is not equal to 0. 
            /// Left-aligned if the iReadingOrder field of the Text record of the 
            /// axis specifies left-to-right reading order; otherwise, right-aligned.
            /// </summary>
            Top = 0x0001,

            /// <summary>
            /// Center-alignment
            /// </summary>
            Center = 0x0002,

            /// <summary>
            /// Bottom-aligned if the trot field of the Text record of the axis is not equal to 0. 
            /// Right-aligned if the iReadingOrder field of the Text record of the 
            /// axis specifies left-to-right reading order; otherwise, left-aligned.
            /// </summary>
            Bottom = 0x0003
        }

        public enum CatLabelType : ushort
        {
            /// <summary>
            /// The value is set to caLabel field as specified by CatSerRange record.
            /// </summary>
            Custom = 0x0000,

            /// <summary>
            /// The value is set to the default value. The number of category (3) labels 
            /// is automatically calculated by the application based on the data in the chart.
            /// </summary>
            Auto = 0x0001
        }

        /// <summary>
        /// An FrtHeaderOld. The frtHeaderOld.rt field MUST be 0x0856.
        /// 
        /// This structure specifies a future record.
        /// </summary>
        public uint frtHeaderOld;

        /// <summary>
        /// An unsigned integer that specifies the distance between the axis and axis label. 
        /// It contains the offset as a percentage of the default distance. 
        /// The default distance is equal to 1/3 the height of the font calculated in pixels. 
        /// 
        /// MUST be a value greater than or equal to 0 (0%) and less than or equal to 1000 (1000%).
        /// </summary>
        public ushort wOffset;

        /// <summary>
        /// An unsigned integer that specifies the alignment of the axis label.
        /// </summary>
        public Alignment at;

        public CatLabelType cAutoCatLabelReal;

        public CatLab(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.frtHeaderOld = reader.ReadUInt32();
            this.wOffset = reader.ReadUInt16();
            this.at = (Alignment)reader.ReadUInt16();
            this.cAutoCatLabelReal = (CatLabelType)Utils.BitmaskToUInt16(reader.ReadUInt16(), 0x0001);

            // ignore last 2 bytes (reserved and optional)
            if (this.Length > 10)
            {
                reader.ReadBytes(2);
            }

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
